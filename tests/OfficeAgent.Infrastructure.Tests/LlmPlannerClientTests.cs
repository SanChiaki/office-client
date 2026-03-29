using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;
using OfficeAgent.Infrastructure.Http;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class LlmPlannerClientTests
    {
        [Fact]
        public void CompletePostsPlannerRequestsToTheConfiguredEndpoint()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"status\":\"completed\","
                    + "\"output\":[{"
                    + "\"type\":\"message\","
                    + "\"role\":\"assistant\","
                    + "\"content\":[{"
                    + "\"type\":\"output_text\","
                    + "\"text\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\"}\""
                    + "}]"
                    + "}]"
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = " https://api.internal.example/ ",
                    Model = "gpt-5-mini",
                });

            var response = client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal("https://api.internal.example/v1/responses", handler.LastRequest.RequestUri.ToString());
            Assert.Equal("Bearer", handler.LastRequest.Headers.Authorization?.Scheme);
            Assert.Equal("secret-token", handler.LastRequest.Headers.Authorization?.Parameter);
            Assert.Contains("Create a summary sheet", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("gpt-5-mini", handler.LastBody, StringComparison.Ordinal);
            Assert.Contains("\"type\":\"json_object\"", handler.LastBody, StringComparison.Ordinal);
            Assert.DoesNotContain("json_schema", handler.LastBody, StringComparison.Ordinal);
            Assert.Equal("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}", response);
        }

        [Fact]
        public void CompletePreservesBaseUrlPathPrefixes()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    "{"
                    + "\"status\":\"completed\","
                    + "\"output\":[{"
                    + "\"type\":\"message\","
                    + "\"role\":\"assistant\","
                    + "\"content\":[{"
                    + "\"type\":\"output_text\","
                    + "\"text\":\"{\\\"mode\\\":\\\"message\\\",\\\"assistantMessage\\\":\\\"ok\\\"}\""
                    + "}]"
                    + "}]"
                    + "}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example/v1/",
                    Model = "gpt-5-mini",
                });

            client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal("https://api.internal.example/v1/responses", handler.LastRequest.RequestUri.ToString());
        }

        [Fact]
        public void CompleteFallsBackToTheLegacyPlannerEndpointWhenResponsesApiIsUnavailable()
        {
            var handler = new RecordingHandler(request =>
            {
                if (request.RequestUri.ToString() == "https://api.internal.example/v1/responses")
                {
                    return new HttpResponseMessage(HttpStatusCode.NotFound)
                    {
                        Content = new StringContent("{\"error\":{\"message\":\"not found\"}}"),
                    };
                }

                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}"),
                };
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            var response = client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            Assert.Equal(2, handler.CallCount);
            Assert.Equal("https://api.internal.example/planner", handler.LastRequest.RequestUri.ToString());
            Assert.Equal("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}", response);
        }

        [Fact]
        public void CompleteRejectsMissingApiKeys()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = " ",
                    BaseUrl = "https://api.internal.example",
                    Model = "gpt-5-mini",
                });

            Action action = () => client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            var error = Assert.Throws<InvalidOperationException>(action);

            Assert.Contains("API Key", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, handler.CallCount);
        }

        [Fact]
        public void CompleteRejectsInvalidBaseUrls()
        {
            var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}"),
            });
            var client = new LlmPlannerClient(
                new HttpClient(handler),
                () => new AppSettings
                {
                    ApiKey = "secret-token",
                    BaseUrl = "api.internal.example",
                    Model = "gpt-5-mini",
                });

            Action action = () => client.Complete(new PlannerRequest
            {
                SessionId = "session-1",
                UserInput = "Create a summary sheet",
            });

            var error = Assert.Throws<InvalidOperationException>(action);

            Assert.Contains("Base URL", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, handler.CallCount);
        }

        private sealed class RecordingHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, HttpResponseMessage> responder;

            public RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
            {
                this.responder = responder;
            }

            public HttpRequestMessage LastRequest { get; private set; }

            public string LastBody { get; private set; } = string.Empty;

            public int CallCount { get; private set; }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                CallCount++;
                LastRequest = request;
                LastBody = request.Content?.ReadAsStringAsync().GetAwaiter().GetResult() ?? string.Empty;
                return Task.FromResult(responder(request));
            }
        }
    }
}
