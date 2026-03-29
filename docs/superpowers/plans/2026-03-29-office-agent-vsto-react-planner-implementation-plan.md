# OfficeAgent Excel VSTO ReAct Planner Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a controlled natural-language ReAct planner to the VSTO MVP so the model can inspect Excel with read steps, produce one frozen write plan, and execute that plan through the native host after a single confirmation.

**Architecture:** Keep the current explicit Excel command path and direct `upload_data` skill path intact. Add a new planner flow in Core that validates model output, executes read-step replanning, freezes write plans, and journals execution. Expose that flow over a new WebView bridge route and render plan preview plus execution status in the React task pane.

**Tech Stack:** C#, .NET Framework 4.8, VSTO, Newtonsoft.Json, HttpClient, React, TypeScript, Vite, Vitest, xUnit

---

## File Structure

- `docs/superpowers/specs/2026-03-29-office-agent-vsto-react-planner-design.md`
  Purpose: approved behavior and protocol for the increment
- `src/OfficeAgent.Core/Models/AgentCommandEnvelope.cs`
  Purpose: extend the agent envelope with planner state and execution payloads
- `src/OfficeAgent.Core/Models/AgentPlan.cs`
  Purpose: new plan, plan step, planner response, journal, and planner constants
- `src/OfficeAgent.Core/Services/IAgentOrchestrator.cs`
  Purpose: keep the orchestrator entrypoint stable while returning richer planner results
- `src/OfficeAgent.Core/Services/ILlmPlannerClient.cs`
  Purpose: planner-facing abstraction for model calls
- `src/OfficeAgent.Core/Services/IPlanExecutor.cs`
  Purpose: execute frozen plan steps through existing command and skill services
- `src/OfficeAgent.Core/Orchestration/AgentOrchestrator.cs`
  Purpose: add controlled ReAct loop, validation, and journaled execution
- `src/OfficeAgent.Core/Services/ConfirmationService.cs`
  Purpose: optionally reuse or extend preview validation for plan steps
- `src/OfficeAgent.Core/Skills/UploadDataSkill.cs`
  Purpose: support plan-step execution without breaking direct skill use
- `src/OfficeAgent.Infrastructure/Http/LlmPlannerClient.cs`
  Purpose: call the configured LLM endpoint and return strict planner JSON
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`
  Purpose: add bridge message types for planner run and plan confirm execution
- `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
  Purpose: route planner requests and confirmations to the orchestrator
- `src/OfficeAgent.Frontend/src/types/bridge.ts`
  Purpose: TypeScript DTOs for planner modes, plans, steps, and journals
- `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
  Purpose: add typed bridge methods for planner run and plan execution
- `src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx`
  Purpose: render both single-command previews and whole-plan previews
- `src/OfficeAgent.Frontend/src/App.tsx`
  Purpose: route natural language to planner, store pending plan confirmations, and render execution results
- `tests/OfficeAgent.Core.Tests/AgentOrchestratorTests.cs`
  Purpose: verify planner loop, validation, freezing, legacy routing, and execution stop conditions
- `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`
  Purpose: verify new planner bridge endpoints and error handling
- `src/OfficeAgent.Frontend/src/App.test.tsx`
  Purpose: verify routing and plan confirmation UX
- `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`
  Purpose: verify new bridge request and response helpers

## Task 1: Define planner models and the failing core tests

**Files:**
- Create: `src/OfficeAgent.Core/Models/AgentPlan.cs`
- Modify: `src/OfficeAgent.Core/Models/AgentCommandEnvelope.cs`
- Modify: `tests/OfficeAgent.Core.Tests/AgentOrchestratorTests.cs`

- [ ] **Step 1: Write the failing planner tests**

```csharp
[Fact]
public void ExecuteReturnsReadStepWhenPlannerRequestsSelectionRead()
{
    var planner = new FakeLlmPlannerClient
    {
        Responses = new[]
        {
            PlannerJson.ReadStep()
        }
    };
    var orchestrator = CreateOrchestrator(planner: planner);

    var result = orchestrator.Execute(new AgentCommandEnvelope
    {
        UserInput = "Summarize the current selection",
        Confirmed = false,
    });

    Assert.Equal(AgentRouteTypes.Plan, result.Route);
    Assert.Equal(PlannerResponseModes.ReadStep, result.Planner.Mode);
}

[Fact]
public void ExecuteReturnsFrozenPlanAfterReadReplanning()
{
    var planner = new FakeLlmPlannerClient
    {
        Responses = new[]
        {
            PlannerJson.ReadStep(),
            PlannerJson.WritePlan()
        }
    };
    var orchestrator = CreateOrchestrator(planner: planner);

    var result = orchestrator.Execute(new AgentCommandEnvelope
    {
        UserInput = "Create a summary sheet with the current selection",
        Confirmed = false,
    });

    Assert.Equal(PlannerResponseModes.Plan, result.Planner.Mode);
    Assert.True(result.RequiresConfirmation);
    Assert.Equal(2, result.Planner.Plan.Steps.Length);
}

[Fact]
public void ExecuteStopsPlanExecutionAfterTheFirstFailedStep()
{
    var orchestrator = CreateOrchestrator(
        planner: new FakeLlmPlannerClient { Responses = new[] { PlannerJson.WritePlan() } },
        executor: new FakePlanExecutor { FailOnStepIndex = 1 });

    var result = orchestrator.Execute(new AgentCommandEnvelope
    {
        UserInput = "Create a summary sheet with the current selection",
        Confirmed = true,
        Plan = PlannerFixtures.CreatePlan(),
    });

    Assert.Equal("failed", result.Status);
    Assert.Equal("completed", result.Journal.Steps[0].Status);
    Assert.Equal("failed", result.Journal.Steps[1].Status);
    Assert.Equal("skipped", result.Journal.Steps[2].Status);
}
```

- [ ] **Step 2: Run the core tests to verify they fail**

Run: `dotnet test tests\OfficeAgent.Core.Tests\OfficeAgent.Core.Tests.csproj --filter AgentOrchestratorTests`

Expected: FAIL with missing planner models, missing planner properties, or missing planner execution behavior.

- [ ] **Step 3: Add planner models with minimal shape**

```csharp
public static class PlannerResponseModes
{
    public const string Message = "message";
    public const string ReadStep = "read_step";
    public const string Plan = "plan";
}

public sealed class PlannerResponse
{
    public string Mode { get; set; } = string.Empty;
    public string AssistantMessage { get; set; } = string.Empty;
    public PlannerReadStep Step { get; set; }
    public AgentPlan Plan { get; set; }
}

public sealed class AgentPlan
{
    public string Summary { get; set; } = string.Empty;
    public AgentPlanStep[] Steps { get; set; } = System.Array.Empty<AgentPlanStep>();
}
```

- [ ] **Step 4: Extend the agent envelope and result types**

```csharp
public sealed class AgentCommandEnvelope
{
    public string UserInput { get; set; } = string.Empty;
    public bool Confirmed { get; set; }
    public AgentPlan Plan { get; set; }
}

public sealed class AgentCommandResult
{
    public string Route { get; set; } = string.Empty;
    public bool RequiresConfirmation { get; set; }
    public PlannerResponse Planner { get; set; }
    public PlanExecutionJournal Journal { get; set; }
}
```

- [ ] **Step 5: Run the same core tests again**

Run: `dotnet test tests\OfficeAgent.Core.Tests\OfficeAgent.Core.Tests.csproj --filter AgentOrchestratorTests`

Expected: FAIL, but now because orchestrator logic is still missing rather than model types being absent.

- [ ] **Step 6: Commit**

```bash
git add src/OfficeAgent.Core/Models/AgentCommandEnvelope.cs src/OfficeAgent.Core/Models/AgentPlan.cs tests/OfficeAgent.Core.Tests/AgentOrchestratorTests.cs
git commit -m "test: define planner models and failing orchestrator coverage"
```

## Task 2: Implement the controlled ReAct loop in Core

**Files:**
- Create: `src/OfficeAgent.Core/Services/ILlmPlannerClient.cs`
- Create: `src/OfficeAgent.Core/Services/IPlanExecutor.cs`
- Modify: `src/OfficeAgent.Core/Orchestration/AgentOrchestrator.cs`
- Modify: `src/OfficeAgent.Core/Skills/UploadDataSkill.cs`
- Modify: `tests/OfficeAgent.Core.Tests/AgentOrchestratorTests.cs`

- [ ] **Step 1: Add failing tests for invalid planner output and loop exhaustion**

```csharp
[Fact]
public void ExecuteReturnsPlannerFailureWhenModelReturnsUnsupportedStepType()
{
    var orchestrator = CreateOrchestrator(
        planner: new FakeLlmPlannerClient { Responses = new[] { PlannerJson.InvalidStep() } });

    var result = orchestrator.Execute(new AgentCommandEnvelope
    {
        UserInput = "Do something unsupported",
        Confirmed = false,
    });

    Assert.Equal(AgentRouteTypes.Chat, result.Route);
    Assert.Equal("failed", result.Status);
    Assert.Contains("supported", result.Message, StringComparison.OrdinalIgnoreCase);
}

[Fact]
public void ExecuteReturnsPlannerFailureWhenLoopBudgetIsExhausted()
{
    var orchestrator = CreateOrchestrator(
        planner: new FakeLlmPlannerClient
        {
            Responses = new[]
            {
                PlannerJson.ReadStep(),
                PlannerJson.ReadStep(),
                PlannerJson.ReadStep(),
            }
        });

    var result = orchestrator.Execute(new AgentCommandEnvelope
    {
        UserInput = "Keep reading forever",
        Confirmed = false,
    });

    Assert.Equal("failed", result.Status);
    Assert.Contains("rephrase", result.Message, StringComparison.OrdinalIgnoreCase);
}
```

- [ ] **Step 2: Run the core tests to verify they fail for the expected reasons**

Run: `dotnet test tests\OfficeAgent.Core.Tests\OfficeAgent.Core.Tests.csproj --filter AgentOrchestratorTests`

Expected: FAIL because `AgentOrchestrator` does not yet parse planner JSON, loop through read steps, or build execution journals.

- [ ] **Step 3: Add the planner client and plan executor interfaces**

```csharp
public interface ILlmPlannerClient
{
    PlannerResponse Plan(PlannerRequest request);
}

public interface IPlanExecutor
{
    PlanExecutionJournal Execute(AgentPlan plan);
}
```

- [ ] **Step 4: Implement minimal planner orchestration**

```csharp
for (var attempt = 0; attempt < 3; attempt++)
{
    var plannerResponse = plannerClient.Plan(request);
    validator.Validate(plannerResponse);

    if (plannerResponse.Mode == PlannerResponseModes.ReadStep)
    {
        var readResult = excelCommandExecutor.Execute(new ExcelCommand
        {
            CommandType = ExcelCommandTypes.ReadSelectionTable,
            Confirmed = false,
        });
        request = request.WithObservation(readResult.Table);
        continue;
    }

    if (plannerResponse.Mode == PlannerResponseModes.Plan)
    {
        return new AgentCommandResult
        {
            Route = AgentRouteTypes.Plan,
            RequiresConfirmation = true,
            Status = "preview",
            Message = plannerResponse.AssistantMessage,
            Planner = plannerResponse,
        };
    }

    return CompleteChat(plannerResponse.AssistantMessage);
}
```

- [ ] **Step 5: Add confirmed-plan execution to the same orchestrator**

```csharp
if (envelope.Confirmed && envelope.Plan != null)
{
    var journal = planExecutor.Execute(envelope.Plan);
    return new AgentCommandResult
    {
        Route = AgentRouteTypes.Plan,
        RequiresConfirmation = false,
        Status = journal.HasFailures ? "failed" : "completed",
        Message = journal.HasFailures ? journal.ErrorMessage : "Plan executed successfully.",
        Journal = journal,
    };
}
```

- [ ] **Step 6: Re-run the orchestrator tests**

Run: `dotnet test tests\OfficeAgent.Core.Tests\OfficeAgent.Core.Tests.csproj --filter AgentOrchestratorTests`

Expected: PASS for the new planner cases while legacy `upload_data` tests remain green.

- [ ] **Step 7: Commit**

```bash
git add src/OfficeAgent.Core/Services/ILlmPlannerClient.cs src/OfficeAgent.Core/Services/IPlanExecutor.cs src/OfficeAgent.Core/Orchestration/AgentOrchestrator.cs src/OfficeAgent.Core/Skills/UploadDataSkill.cs tests/OfficeAgent.Core.Tests/AgentOrchestratorTests.cs
git commit -m "feat: add controlled react planner orchestration"
```

## Task 3: Implement native plan execution and LLM infrastructure

**Files:**
- Create: `src/OfficeAgent.Infrastructure/Http/LlmPlannerClient.cs`
- Modify: `src/OfficeAgent.Infrastructure/Http/BusinessApiClient.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`
- Modify: `tests/OfficeAgent.Infrastructure.Tests/BusinessApiClientTests.cs`
- Create: `tests/OfficeAgent.Infrastructure.Tests/LlmPlannerClientTests.cs`

- [ ] **Step 1: Write failing infrastructure tests for planner HTTP calls**

```csharp
[Fact]
public void PlanPostsPlannerRequestToConfiguredChatEndpoint()
{
    var handler = new RecordingHandler("{\"mode\":\"message\",\"assistantMessage\":\"ok\"}");
    var client = new LlmPlannerClient(
        new HttpClient(handler) { BaseAddress = new Uri("https://api.example.com/") },
        () => new AppSettings { ApiKey = "secret", BaseUrl = "https://api.example.com", Model = "gpt-5-mini" });

    var response = client.Plan(new PlannerRequest { UserInput = "hello" });

    Assert.Equal(HttpMethod.Post, handler.LastRequest.Method);
    Assert.Equal("/planner", handler.LastRequest.RequestUri.AbsolutePath);
    Assert.Equal("message", response.Mode);
}
```

- [ ] **Step 2: Run the infrastructure tests and watch them fail**

Run: `dotnet test tests\OfficeAgent.Infrastructure.Tests\OfficeAgent.Infrastructure.Tests.csproj --filter LlmPlannerClientTests`

Expected: FAIL because `LlmPlannerClient` does not exist yet.

- [ ] **Step 3: Implement the minimal planner client**

```csharp
public sealed class LlmPlannerClient : ILlmPlannerClient
{
    public PlannerResponse Plan(PlannerRequest request)
    {
        var payload = JsonConvert.SerializeObject(new
        {
            model = settings.Model,
            input = plannerPromptBuilder.Build(request)
        });

        using (var httpRequest = new HttpRequestMessage(HttpMethod.Post, new Uri(baseUri, "/planner")))
        {
            httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
            httpRequest.Content = new StringContent(payload, Encoding.UTF8, "application/json");
            var response = httpClient.SendAsync(httpRequest).GetAwaiter().GetResult();
            var responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            return JsonConvert.DeserializeObject<PlannerResponse>(responseBody);
        }
    }
}
```

- [ ] **Step 4: Register the planner client in the add-in bootstrap**

```csharp
var plannerClient = new LlmPlannerClient(new HttpClient(), settingsStore.Load);
var orchestrator = new AgentOrchestrator(skillRegistry, excelCommandExecutor, plannerClient, planExecutor);
```

- [ ] **Step 5: Re-run the infrastructure tests**

Run: `dotnet test tests\OfficeAgent.Infrastructure.Tests\OfficeAgent.Infrastructure.Tests.csproj --filter "LlmPlannerClientTests|BusinessApiClientTests"`

Expected: PASS for the planner client coverage and no regressions in business API tests.

- [ ] **Step 6: Commit**

```bash
git add src/OfficeAgent.Infrastructure/Http/LlmPlannerClient.cs src/OfficeAgent.ExcelAddIn/ThisAddIn.cs tests/OfficeAgent.Infrastructure.Tests/LlmPlannerClientTests.cs tests/OfficeAgent.Infrastructure.Tests/BusinessApiClientTests.cs
git commit -m "feat: add planner http client and native wiring"
```

## Task 4: Expose planner run and plan execution over the bridge

**Files:**
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs`
- Modify: `src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs`
- Modify: `tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs`

- [ ] **Step 1: Write failing bridge tests**

```csharp
[Fact]
public void RunAgentRoutesNaturalLanguageRequestsThroughThePlanner()
{
    var orchestrator = new FakeAgentOrchestrator
    {
        Result = AgentResults.PlanPreview()
    };
    var router = CreateRouter(..., agentOrchestrator: orchestrator);

    var responseJson = InvokeRoute(
        router,
        "{\"type\":\"bridge.runAgent\",\"requestId\":\"req-1\",\"payload\":{\"userInput\":\"Create a summary sheet\",\"confirmed\":false}}");

    Assert.Contains("\"ok\":true", responseJson);
    Assert.Contains("\"mode\":\"plan\"", responseJson);
}

[Fact]
public void RunAgentRejectsMissingPayload()
{
    var router = CreateRouter(...);
    var responseJson = InvokeRoute(
        router,
        "{\"type\":\"bridge.runAgent\",\"requestId\":\"req-1\"}");

    Assert.Contains("\"ok\":false", responseJson);
    Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
}
```

- [ ] **Step 2: Run the bridge tests to verify they fail**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter WebMessageRouterTests`

Expected: FAIL because `bridge.runAgent` is not a known message type.

- [ ] **Step 3: Add the new message type and route**

```csharp
internal static class BridgeMessageTypes
{
    public const string RunAgent = "bridge.runAgent";
}

case BridgeMessageTypes.RunAgent:
    return RunAgent(request);
```

- [ ] **Step 4: Implement `RunAgent` payload parsing**

```csharp
private WebMessageResponse RunAgent(WebMessageRequest request)
{
    var envelope = request.Payload.ToObject<AgentCommandEnvelope>() ?? new AgentCommandEnvelope();
    return Success(request.Type, request.RequestId, agentOrchestrator.Execute(envelope));
}
```

- [ ] **Step 5: Re-run the bridge tests**

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj --filter WebMessageRouterTests`

Expected: PASS and no regressions in existing `bridge.runSkill` or `bridge.executeExcelCommand` coverage.

- [ ] **Step 6: Commit**

```bash
git add src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageEnvelope.cs src/OfficeAgent.ExcelAddIn/WebBridge/WebMessageRouter.cs tests/OfficeAgent.ExcelAddIn.Tests/WebMessageRouterTests.cs
git commit -m "feat: expose planner flow over the web bridge"
```

## Task 5: Add planner types and UX to the React task pane

**Files:**
- Modify: `src/OfficeAgent.Frontend/src/types/bridge.ts`
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts`
- Modify: `src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx`
- Modify: `src/OfficeAgent.Frontend/src/App.tsx`
- Modify: `src/OfficeAgent.Frontend/src/App.test.tsx`
- Modify: `src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts`

- [ ] **Step 1: Write failing frontend tests for natural-language planning**

```tsx
it('routes plain natural language to the planner bridge', async () => {
  render(<App />);

  await user.type(screen.getByLabelText('Message composer'), 'Create a summary sheet from the current selection');
  await user.click(screen.getByRole('button', { name: 'Send' }));

  expect(nativeBridge.runAgent).toHaveBeenCalledWith({
    userInput: 'Create a summary sheet from the current selection',
    confirmed: false,
  });
});

it('renders a plan preview and confirms the frozen plan', async () => {
  nativeBridge.runAgent.mockResolvedValue(bridgeFixtures.planPreview());
  render(<App />);

  await user.type(screen.getByLabelText('Message composer'), 'Create a summary sheet from the current selection');
  await user.click(screen.getByRole('button', { name: 'Send' }));

  expect(await screen.findByText('Create a Summary sheet and write the selected rows.')).toBeInTheDocument();
  await user.click(screen.getByRole('button', { name: 'Confirm' }));

  expect(nativeBridge.runAgent).toHaveBeenLastCalledWith({
    userInput: 'Create a summary sheet from the current selection',
    confirmed: true,
    plan: bridgeFixtures.planPreview().planner.plan,
  });
});
```

- [ ] **Step 2: Run the frontend tests and watch them fail**

Run: `npm.cmd test -- src/App.test.tsx src/bridge/nativeBridge.test.ts`

Expected: FAIL because `runAgent` and planner DTOs do not exist yet.

- [ ] **Step 3: Add planner DTOs and bridge methods**

```ts
export interface PlannerResponse {
  mode: 'message' | 'read_step' | 'plan';
  assistantMessage: string;
  plan?: AgentPlan;
}

export interface AgentResult {
  route: string;
  status: string;
  requiresConfirmation: boolean;
  planner?: PlannerResponse;
  journal?: PlanExecutionJournal;
}

async function runAgent(payload: AgentRequestEnvelope): Promise<AgentResult> {
  return sendRequest<AgentResult>('bridge.runAgent', payload);
}
```

- [ ] **Step 4: Update the composer flow and pending confirmation state**

```tsx
if (command) {
  await dispatchExcelCommand(command, sessionId);
  return;
}

if (trimmedValue.startsWith('/upload_data')) {
  await dispatchSkill({ userInput: trimmedValue, confirmed: false }, sessionId);
  return;
}

await dispatchAgent({ userInput: trimmedValue, confirmed: false }, sessionId);
```

- [ ] **Step 5: Update the confirmation card to render plan steps**

```tsx
{preview.kind === 'plan' ? (
  <ol>
    {preview.plan.steps.map((step, index) => (
      <li key={`${step.type}-${index}`}>{formatPlanStep(step)}</li>
    ))}
  </ol>
) : (
  <ul>
    {preview.details.map((detail) => <li key={detail}>{detail}</li>)}
  </ul>
)}
```

- [ ] **Step 6: Re-run the frontend tests**

Run: `npm.cmd test -- src/App.test.tsx src/bridge/nativeBridge.test.ts`

Expected: PASS for the new planner routing and preview behavior, with existing command and skill flows still green.

- [ ] **Step 7: Commit**

```bash
git add src/OfficeAgent.Frontend/src/types/bridge.ts src/OfficeAgent.Frontend/src/bridge/nativeBridge.ts src/OfficeAgent.Frontend/src/components/ConfirmationCard.tsx src/OfficeAgent.Frontend/src/App.tsx src/OfficeAgent.Frontend/src/App.test.tsx src/OfficeAgent.Frontend/src/bridge/nativeBridge.test.ts
git commit -m "feat: add planner flow to the task pane ui"
```

## Task 6: Full verification and cleanup

**Files:**
- Modify: `docs/vsto-manual-test-checklist.md`
- Modify: `docs/vsto-known-limitations.md`

- [ ] **Step 1: Update the manual QA checklist**

```markdown
- natural language request with no slash triggers planner
- planner can read selection before presenting a plan
- write plan appears once with ordered steps
- confirm executes steps in order
- first failed step stops the remaining steps
- explicit slash Excel commands still work
- /upload_data still bypasses planner and keeps its original confirmation flow
```

- [ ] **Step 2: Run the full automated test suite**

Run: `dotnet test tests\OfficeAgent.Core.Tests\OfficeAgent.Core.Tests.csproj`

Expected: PASS

Run: `dotnet test tests\OfficeAgent.Infrastructure.Tests\OfficeAgent.Infrastructure.Tests.csproj`

Expected: PASS

Run: `dotnet test tests\OfficeAgent.ExcelAddIn.Tests\OfficeAgent.ExcelAddIn.Tests.csproj`

Expected: PASS

Run: `npm.cmd test`

Expected: PASS

Run: `npm.cmd run build`

Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add docs/vsto-manual-test-checklist.md docs/vsto-known-limitations.md
git commit -m "chore: document planner verification coverage"
```
