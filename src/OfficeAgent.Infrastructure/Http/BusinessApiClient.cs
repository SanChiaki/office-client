using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.Infrastructure.Http
{
    public sealed class BusinessApiClient : IUploadDataGateway
    {
        private readonly HttpClient httpClient;
        private readonly Func<AppSettings> loadSettings;

        public BusinessApiClient(FileSettingsStore settingsStore, HttpClient httpClient = null)
            : this(() => settingsStore?.Load() ?? new AppSettings(), httpClient)
        {
        }

        public BusinessApiClient(HttpClient httpClient, Func<AppSettings> loadSettings)
            : this(loadSettings, httpClient)
        {
        }

        public BusinessApiClient(Func<AppSettings> loadSettings, HttpClient httpClient = null)
        {
            this.loadSettings = loadSettings ?? throw new ArgumentNullException(nameof(loadSettings));
            this.httpClient = httpClient ?? new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(15),
            };
        }

        public UploadExecutionResult Upload(UploadPreview preview)
        {
            if (preview == null)
            {
                throw new ArgumentNullException(nameof(preview));
            }

            var settings = loadSettings() ?? new AppSettings();
            var baseUrl = AppSettings.NormalizeBaseUrl(settings.BaseUrl);
            if (string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                throw new InvalidOperationException("An API Key is required before upload_data can call the business API.");
            }

            if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) ||
                (baseUri.Scheme != Uri.UriSchemeHttp && baseUri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("The configured Business API Base URL is invalid. Update settings and try again.");
            }

            var endpoint = new Uri($"{baseUri.AbsoluteUri.TrimEnd('/')}/upload_data");
            var payload = JsonConvert.SerializeObject(new
            {
                projectName = preview.ProjectName,
                source = new
                {
                    sheetName = preview.SheetName,
                    address = preview.Address,
                },
                headers = preview.Headers,
                records = preview.Records,
            });

            for (var attempt = 1; attempt <= 2; attempt++)
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                    request.Content = new StringContent(payload, Encoding.UTF8, "application/json");

                    try
                    {
                        using (var response = httpClient.SendAsync(request).GetAwaiter().GetResult())
                        {
                            if (!response.IsSuccessStatusCode)
                            {
                                var errorBody = response.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                                if (attempt < 2 && IsRetryableStatusCode(response.StatusCode))
                                {
                                    continue;
                                }

                                throw new InvalidOperationException(
                                    FormatErrorMessage(response.StatusCode, response.ReasonPhrase, errorBody));
                            }

                            var responseBody = response.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                            var parsed = TryParse(responseBody);
                            return new UploadExecutionResult
                            {
                                SavedCount = parsed?.SavedCount ?? preview.Records.Length,
                                Message = !string.IsNullOrWhiteSpace(parsed?.Message)
                                    ? parsed.Message
                                    : $"Uploaded {preview.Records.Length} row(s) to {preview.ProjectName}.",
                            };
                        }
                    }
                    catch (HttpRequestException error)
                    {
                        if (attempt >= 2)
                        {
                            throw new InvalidOperationException($"Business API request failed: {error.Message}", error);
                        }
                    }
                    catch (OperationCanceledException error)
                    {
                        throw new InvalidOperationException(
                            $"Business API request timed out after {httpClient.Timeout.TotalSeconds:0} seconds.",
                            error);
                    }
                }
            }

            throw new InvalidOperationException("Business API request failed after retrying.");
        }

        private static bool IsRetryableStatusCode(HttpStatusCode statusCode)
        {
            var numericStatusCode = (int)statusCode;
            return numericStatusCode >= 500 || statusCode == HttpStatusCode.RequestTimeout;
        }

        private static UploadExecutionResponse TryParse(string responseBody)
        {
            if (string.IsNullOrWhiteSpace(responseBody))
            {
                return null;
            }

            try
            {
                return JsonConvert.DeserializeObject<UploadExecutionResponse>(responseBody);
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static string FormatErrorMessage(HttpStatusCode statusCode, string reasonPhrase, string responseBody)
        {
            var parsedError = TryParseError(responseBody);
            if (parsedError != null &&
                (!string.IsNullOrWhiteSpace(parsedError.Code) || !string.IsNullOrWhiteSpace(parsedError.Message)))
            {
                var formattedCode = string.IsNullOrWhiteSpace(parsedError.Code) ? reasonPhrase : parsedError.Code;
                var formattedMessage = string.IsNullOrWhiteSpace(parsedError.Message) ? responseBody : parsedError.Message;
                return $"Business API request failed ({(int)statusCode} {formattedCode}): {formattedMessage}";
            }

            var responseSummary = string.IsNullOrWhiteSpace(responseBody)
                ? "No response body was returned."
                : responseBody;
            return $"Business API request failed ({(int)statusCode} {reasonPhrase}): {responseSummary}";
        }

        private static BusinessApiErrorResponse TryParseError(string responseBody)
        {
            if (string.IsNullOrWhiteSpace(responseBody))
            {
                return null;
            }

            try
            {
                return JsonConvert.DeserializeObject<BusinessApiErrorResponse>(responseBody);
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private sealed class UploadExecutionResponse
        {
            [JsonProperty("savedCount")]
            public int SavedCount { get; set; }

            [JsonProperty("message")]
            public string Message { get; set; } = string.Empty;
        }

        private sealed class BusinessApiErrorResponse
        {
            [JsonProperty("code")]
            public string Code { get; set; } = string.Empty;

            [JsonProperty("message")]
            public string Message { get; set; } = string.Empty;
        }
    }
}
