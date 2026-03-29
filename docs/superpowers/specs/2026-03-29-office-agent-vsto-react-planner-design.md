# OfficeAgent Excel VSTO ReAct Planner Design

Date: 2026-03-29

Status: Approved for implementation

## 1. Goal

Add a controlled natural-language agent flow to the existing VSTO MVP so the model can:

- inspect the current Excel context through allowed read actions
- decide which supported Excel actions or skills to use
- decide the action order
- present one frozen write plan for confirmation
- execute the confirmed plan step-by-step through the native host

This is an incremental change on top of the current VSTO MVP. Explicit Excel commands and the existing `upload_data` skill must keep working.

## 2. Scope

This increment only exposes the following actions to the model:

- `excel.readSelectionTable`
- `excel.writeRange`
- `excel.addWorksheet`
- `excel.renameWorksheet`
- `excel.deleteWorksheet`
- `skill.upload_data`

The model does not get arbitrary code execution, arbitrary Excel automation, formula synthesis, formatting actions, row or column structural edits, or unrestricted external API access.

## 3. Product Behavior

### 3.1 User interaction model

When the user enters a message that is not an explicit slash Excel command and is not a direct slash skill invocation, OfficeAgent routes the message to the planner.

The planner uses a controlled ReAct loop:

1. the host sends the user request, session context, current selection summary, and allowed action catalog to the model
2. the model may request one read step at a time through `excel.readSelectionTable`
3. the host executes the read step immediately and feeds the result back to the model
4. once the model decides that a write or side-effect is needed, it must stop iterating and return a complete frozen execution plan
5. the UI shows that plan once for user confirmation
6. after confirmation, the host executes the plan in order without asking the model to revise it

### 3.2 Confirmation policy

The planning loop may execute read actions immediately.

The following actions are always treated as plan actions that require one plan-level confirmation:

- `excel.writeRange`
- `excel.addWorksheet`
- `excel.renameWorksheet`
- `excel.deleteWorksheet`
- `skill.upload_data`

There is no per-step confirmation inside the plan executor for this increment. If any step fails, execution stops and the remaining steps are marked as not run.

### 3.3 Context sent to the model

The default planning context includes only:

- user input
- current session id and a compact message summary
- current selection metadata
- selection headers
- first few sample rows
- supported actions and their JSON argument shapes

The host must not send the full selected table by default. If the model needs more than the summary, it must request `excel.readSelectionTable`.

## 4. Planner Protocol

The model must return one of three structured response modes:

### 4.1 `message`

Used when the assistant only needs to answer the user and does not need to act.

Example shape:

```json
{
  "mode": "message",
  "assistantMessage": "I cannot safely complete that with the currently supported Excel actions."
}
```

### 4.2 `read_step`

Used when the assistant needs more read-only Excel context before forming a plan.

Only this step type is allowed in this mode for the current increment:

- `excel.readSelectionTable`

Example shape:

```json
{
  "mode": "read_step",
  "assistantMessage": "I will read the current selection before preparing the write plan.",
  "step": {
    "type": "excel.readSelectionTable",
    "args": {}
  }
}
```

### 4.3 `plan`

Used when the assistant is ready to propose a full action plan that contains every remaining write or side-effect step in execution order.

Example shape:

```json
{
  "mode": "plan",
  "assistantMessage": "I prepared a plan. Review it before Excel is changed.",
  "plan": {
    "summary": "Create a Summary sheet and write the selected table into it.",
    "steps": [
      {
        "type": "excel.addWorksheet",
        "args": {
          "newSheetName": "Summary"
        }
      },
      {
        "type": "excel.writeRange",
        "args": {
          "targetAddress": "Summary!A1:B3",
          "values": [
            ["Name", "Region"],
            ["Project A", "CN"],
            ["Project B", "US"]
          ]
        }
      }
    ]
  }
}
```

## 5. Native State Machine

The native planner service follows this state machine:

- `Idle`
  Wait for a user prompt.
- `Planning`
  Call the model with current context.
- `Reading`
  Execute one read step returned by the model.
- `Replanning`
  Append the read result to planner history and call the model again.
- `PlanReady`
  Store the frozen plan and return it to the frontend for confirmation.
- `Executing`
  Execute each planned step in order.
- `Completed`
  Persist final journal entries and assistant summary.
- `Failed`
  Persist the failure point, completed steps, skipped steps, and error message.

The planner loop is capped at three model turns for the initial increment. If the limit is reached without a valid `message` or `plan` response, the host returns a user-facing error and logs the invalid loop outcome.

## 6. Validation Rules

Before any model response is accepted, the host validates it.

### 6.1 Global rules

- the response must include a supported `mode`
- `assistantMessage` must be present and non-empty
- unknown top-level fields are ignored for now

### 6.2 `read_step` rules

- exactly one step is allowed
- the step type must be `excel.readSelectionTable`
- arguments must be empty or omitted

### 6.3 `plan` rules

- `plan.summary` must be present
- `plan.steps` must contain at least one step
- every step type must be one of the allowed plan actions
- every step must have arguments that can be converted into an existing native command or skill envelope
- plan steps may not contain `excel.readSelectionTable`

### 6.4 Plan safety rules

- `excel.writeRange` must still pass the existing guard and confirmation validation
- `skill.upload_data` must still produce an `UploadPreview` before the actual upload call is committed
- invalid or ambiguous plans are rejected before the UI sees them

## 7. Execution Model

### 7.1 Read loop execution

When the planner returns `read_step`, the host immediately executes `excel.readSelectionTable` through the existing `IExcelCommandExecutor`.

The resulting `ExcelTableData` is normalized into a compact planner observation payload and appended to the model conversation.

### 7.2 Frozen plan execution

Once the planner returns `plan`, the plan is frozen and stored with the session-level pending confirmation state.

After the user confirms:

- the host executes each step in order
- each Excel step uses the existing native executor
- each skill step uses the existing orchestrator or skill registry path
- the executor records a journal entry per step
- execution stops on the first failure

The model is not called again during plan execution for this increment.

## 8. Frontend Changes

The frontend needs three additive changes:

### 8.1 New agent dispatch path

The composer must route plain natural-language prompts to a new bridge call instead of always calling `runSkill`.

Direct slash Excel commands remain parsed locally first. Direct slash skill commands still bypass the planner and continue to use the existing skill path.

### 8.2 Plan preview card

The confirmation card must support a plan preview with:

- plan summary
- ordered steps
- clear action labels
- plan-level confirm and cancel actions

### 8.3 Execution journal rendering

After execution, the thread must show:

- assistant planning message
- plan summary
- per-step completion or failure status
- final assistant completion or failure message

## 9. Data Model Additions

The system needs explicit planner models in both .NET and TypeScript:

- planner mode enum or constants
- planner response DTO
- read-step DTO
- execution plan DTO
- execution plan step DTO
- execution journal DTO
- pending confirmation state for a plan

These models must stay bridge-safe and JSON-serializable.

## 10. Prompting Strategy

The planner prompt must tell the model:

- it is not allowed to invent unsupported actions
- it must return strict JSON only
- it may use `read_step` only for `excel.readSelectionTable`
- it must not emit a write step directly outside `plan`
- it must produce a complete ordered plan before any write or side effect occurs
- it should prefer the smallest safe plan that satisfies the request

The prompt must also include compact action documentation and example argument shapes for each supported action.

## 11. Failure Handling

Failure cases and required behavior:

- malformed model JSON
  Return a planner error message and log the raw response for diagnostics.
- unsupported planner mode or step type
  Reject the response and fail the planner turn.
- planner loop exhaustion
  Return a stable user-facing message asking the user to rephrase or use an explicit command.
- read step failure
  Stop planning and surface the native error.
- plan validation failure
  Return a stable user-facing message and log the invalid plan.
- execution step failure
  Stop execution immediately and report which steps completed and which did not run.

## 12. Testing Strategy

The increment requires test coverage in three layers:

### 12.1 Core tests

- planner response validation
- planner loop transitions
- plan freezing
- stop-on-first-failure execution behavior
- fallback handling for invalid model output
- legacy `upload_data` route still working

### 12.2 Add-in bridge tests

- new bridge message type for agent planning
- malformed payload protection
- successful plan response serialization
- successful plan execution serialization
- unexpected planner exceptions returning `internal_error`

### 12.3 Frontend tests

- natural language input routes to the planner bridge
- slash Excel commands still route locally
- slash skill commands still use the skill bridge
- plan preview card renders ordered steps
- confirm triggers the frozen plan execution path
- failure journal renders correctly

## 13. Out of Scope

This increment does not add:

- autonomous multi-turn execution after write steps start
- dynamic replanning during execution
- arbitrary Excel formulas or formatting
- row and column insert or delete operations
- cloud session sync
- OAuth
- prompt streaming
- parallel plan step execution
