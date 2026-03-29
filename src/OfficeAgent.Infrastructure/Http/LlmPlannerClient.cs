using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class LlmPlannerClient : ILlmPlannerClient
    {
        private readonly HttpClient httpClient;
        private readonly Func<AppSettings> loadSettings;

        public LlmPlannerClient(FileSettingsStore settingsStore, HttpClient httpClient = null)
            : this(httpClient, () => settingsStore?.Load() ?? new AppSettings())
        {
        }

        public LlmPlannerClient(HttpClient httpClient, Func<AppSettings> loadSettings)
        {
            this.httpClient = httpClient ?? new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(30),
            };
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
        }

        public string Complete(PlannerRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            var settings = loadSettings() ?? new AppSettings();
            var baseUrl = AppSettings.NormalizeBaseUrl(settings.BaseUrl);
            if (string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                throw new InvalidOperationException("An API Key is required before agent planning can call the planner API.");
            }

            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured Planner API Base URL is invalid. Update settings and try again.");
            }

            try
            {
                return CompleteWithOpenAiCompatibleResponses(baseUri, settings, request);
            }
            catch (LegacyPlannerFallbackException)
            {
                return CompleteWithLegacyPlanner(baseUri, settings, request);
            }
        }

        private string CompleteWithOpenAiCompatibleResponses(Uri baseUri, AppSettings settings, PlannerRequest request)
        {
            var endpoint = BuildResponsesEndpoint(baseUri);
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                input = new object[]
                {
                    CreateTextMessage("system", BuildPlannerInstructions()),
                    CreateTextMessage("user", BuildPlannerPrompt(request)),
                },
                text = new
                {
                    format = new
                    {
                        type = "json_object",
                    },
                },
            });

            var responseBody = SendRequest(endpoint, settings.ApiKey, payload, allowLegacyFallback: true);
            return ExtractResponsesText(responseBody);
        }

        private string CompleteWithLegacyPlanner(Uri baseUri, AppSettings settings, PlannerRequest request)
        {
            var endpoint = new Uri($"{baseUri.AbsoluteUri.TrimEnd('/')}/planner");
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                request,
            });
            return SendRequest(endpoint, settings.ApiKey, payload, allowLegacyFallback: false);
        }

        private string SendRequest(Uri endpoint, string apiKey, string payload, bool allowLegacyFallback)
        {
            using (var httpRequest = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                httpRequest.Content = new StringContent(payload, Encoding.UTF8, "application/json");

                using (var response = httpClient.SendAsync(httpRequest).GetAwaiter().GetResult())
                {
                    var responseBody = response.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                    if (!response.IsSuccessStatusCode)
                    {
                        if (allowLegacyFallback &&
                            ((int)response.StatusCode == 404 || (int)response.StatusCode == 405))
                        {
                            throw new LegacyPlannerFallbackException();
                        }

                        throw new InvalidOperationException(
                            $"Planner API request failed ({(int)response.StatusCode} {response.ReasonPhrase}): {responseBody}");
                    }

                    if (string.IsNullOrWhiteSpace(responseBody))
                    {
                        throw new InvalidOperationException("Planner API returned an empty response body.");
                    }

                    return responseBody;
                }
            }
        }

        private static Uri BuildResponsesEndpoint(Uri baseUri)
        {
            var absoluteUri = baseUri.AbsoluteUri.TrimEnd('/');
            var absolutePath = baseUri.AbsolutePath?.Trim('/') ?? string.Empty;
            if (string.IsNullOrWhiteSpace(absolutePath))
            {
                return new Uri($"{absoluteUri}/v1/responses");
            }

            return new Uri($"{absoluteUri}/responses");
        }

        private static object CreateTextMessage(string role, string text)
        {
            return new
            {
                role,
                content = new object[]
                {
                    new
                    {
                        type = "input_text",
                        text,
                    },
                },
            };
        }

        private static string BuildPlannerPrompt(PlannerRequest request)
        {
            return "Planner request:\n" + JsonConvert.SerializeObject(
                request ?? new PlannerRequest(),
                Formatting.Indented,
                new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                });
        }

        private static string BuildPlannerInstructions()
        {
            return "You are OfficeAgent's planner. "
                + "Return exactly one JSON object and no markdown. "
                + "Always include the keys mode, assistantMessage, step, and plan. "
                + "Use null for step or plan when they do not apply. "
                + "assistantMessage should be concise and use the user's language when possible. "
                + "Supported modes are message, read_step, and plan. "
                + "Use message when no Excel action is needed or the request is unsupported. "
                + "Use read_step only when you need the full current selection table before planning. "
                + "The only supported read_step is excel.readSelectionTable with empty args. "
                + "Use plan for any write or side-effect sequence. "
                + "Supported plan step types are excel.writeRange, excel.addWorksheet, excel.renameWorksheet, excel.deleteWorksheet, and skill.upload_data. "
                + "Never invent other step types. "
                + "For excel.writeRange use args targetAddress and values. "
                + "For excel.addWorksheet use arg newSheetName. "
                + "For excel.renameWorksheet use args sheetName and newSheetName. "
                + "For excel.deleteWorksheet use arg sheetName. "
                + "For skill.upload_data use arg userInput and preserve the user's upload intent. "
                + "Use the provided selection metadata, headers, sample rows, and prior observations. "
                + "Only request read_step when the summary is insufficient. "
                + "When mode=read_step, set step to {\"type\":\"excel.readSelectionTable\",\"args\":{}} and plan to null. "
                + "When mode=plan, set plan.summary and plan.steps, and set step to null. "
                + "When mode=message, set both step and plan to null. "
                + "If the request cannot be completed safely with the supported actions, answer with mode=message.";
        }

        private static string ExtractResponsesText(string responseBody)
        {
            try
            {
                var parsed = JObject.Parse(responseBody);
                var outputText = parsed["output_text"]?.Value<string>();
                if (!string.IsNullOrWhiteSpace(outputText))
                {
                    return outputText;
                }

                var outputItems = parsed["output"] as JArray;
                if (outputItems != null)
                {
                    foreach (var outputItem in outputItems)
                    {
                        var contentItems = outputItem["content"] as JArray;
                        if (contentItems == null)
                        {
                            continue;
                        }

                        foreach (var contentItem in contentItems)
                        {
                            if (string.Equals(contentItem["type"]?.Value<string>(), "output_text", StringComparison.Ordinal))
                            {
                                var contentText = contentItem["text"]?.Value<string>();
                                if (!string.IsNullOrWhiteSpace(contentText))
                                {
                                    return contentText;
                                }
                            }
                        }
                    }
                }
            }
            catch (JsonException)
            {
                throw new InvalidOperationException("Planner API returned a non-JSON Responses payload.");
            }

            throw new InvalidOperationException("Planner API returned a Responses payload without planner text output.");
        }

        private sealed class LegacyPlannerFallbackException : Exception
        {
        }
    }
}
