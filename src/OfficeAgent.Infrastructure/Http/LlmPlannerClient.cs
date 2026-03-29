using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
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

            var endpoint = new Uri($"{baseUri.AbsoluteUri.TrimEnd('/')}/planner");
            var payload = JsonConvert.SerializeObject(new
            {
                model = settings.Model,
                request,
            });

            using (var httpRequest = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
                httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                httpRequest.Content = new StringContent(payload, Encoding.UTF8, "application/json");

                using (var response = httpClient.SendAsync(httpRequest).GetAwaiter().GetResult())
                {
                    var responseBody = response.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                    if (!response.IsSuccessStatusCode)
                    {
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
    }
}
