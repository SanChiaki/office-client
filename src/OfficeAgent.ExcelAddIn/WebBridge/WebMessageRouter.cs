using System;
using System.Collections.Generic;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeAgent.ExcelAddIn.WebBridge
{
    internal sealed class WebMessageRouter
    {
        private readonly HashSet<string> allowedTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            BridgeMessageTypes.Ping,
            BridgeMessageTypes.GetSelectionContext,
            BridgeMessageTypes.GetSessions,
            BridgeMessageTypes.SaveSettings,
            BridgeMessageTypes.ExecuteExcelCommand,
            BridgeMessageTypes.RunSkill,
        };

        public string Route(string rawRequestJson)
        {
            var response = RouteInternal(rawRequestJson);
            return JsonConvert.SerializeObject(response);
        }

        private WebMessageResponse RouteInternal(string rawRequestJson)
        {
            WebMessageRequest request;

            try
            {
                request = JsonConvert.DeserializeObject<WebMessageRequest>(rawRequestJson);
            }
            catch (JsonException)
            {
                return Error(
                    type: "bridge.unknown",
                    requestId: string.Empty,
                    code: "malformed_json",
                    message: "The web message payload was not valid JSON.");
            }

            if (request == null ||
                string.IsNullOrWhiteSpace(request.Type) ||
                string.IsNullOrWhiteSpace(request.RequestId))
            {
                return Error(
                    type: request?.Type ?? "bridge.unknown",
                    requestId: request?.RequestId ?? string.Empty,
                    code: "malformed_request",
                    message: "Web messages must include both type and requestId.");
            }

            if (!allowedTypes.Contains(request.Type))
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "unknown_message",
                    message: $"Message type '{request.Type}' is not allowed.");
            }

            switch (request.Type)
            {
                case BridgeMessageTypes.Ping:
                    if (HasUnexpectedPayload(request.Payload))
                    {
                        return Error(
                            request.Type,
                            request.RequestId,
                            code: "malformed_payload",
                            message: "bridge.ping does not accept a payload.");
                    }

                    return Success(
                        request.Type,
                        request.RequestId,
                        new PingPayload
                        {
                            Host = "OfficeAgent.ExcelAddIn",
                            Version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "dev",
                        });
                case BridgeMessageTypes.GetSelectionContext:
                case BridgeMessageTypes.GetSessions:
                case BridgeMessageTypes.SaveSettings:
                case BridgeMessageTypes.ExecuteExcelCommand:
                case BridgeMessageTypes.RunSkill:
                    return Error(
                        request.Type,
                        request.RequestId,
                        code: "not_implemented",
                        message: $"Message type '{request.Type}' is registered but not implemented yet.");
                default:
                    return Error(
                        request.Type,
                        request.RequestId,
                        code: "unknown_message",
                        message: $"Message type '{request.Type}' is not allowed.");
            }
        }

        private static bool HasUnexpectedPayload(JToken payload)
        {
            if (payload == null || payload.Type == JTokenType.Null)
            {
                return false;
            }

            return payload.Type != JTokenType.Object || payload.HasValues;
        }

        private static WebMessageResponse Success(string type, string requestId, object payload)
        {
            return new WebMessageResponse
            {
                Type = type,
                RequestId = requestId,
                Ok = true,
                Payload = payload,
            };
        }

        private static WebMessageResponse Error(string type, string requestId, string code, string message)
        {
            return new WebMessageResponse
            {
                Type = type,
                RequestId = requestId,
                Ok = false,
                Error = new WebMessageError
                {
                    Code = code,
                    Message = message,
                },
            };
        }
    }
}
