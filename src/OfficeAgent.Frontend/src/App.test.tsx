import { cleanup, render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import App from './App';
import { nativeBridge } from './bridge/nativeBridge';
import type { SelectionContext } from './types/bridge';

vi.mock('./bridge/nativeBridge', () => ({
  nativeBridge: {
    ping: vi.fn(),
    getSelectionContext: vi.fn(),
    getSessions: vi.fn(),
    onSelectionContextChanged: vi.fn(),
    getSettings: vi.fn(),
    saveSettings: vi.fn(),
    executeExcelCommand: vi.fn(),
    runSkill: vi.fn(),
    runAgent: vi.fn(),
  },
}));

const mockedBridge = vi.mocked(nativeBridge);
let selectionContextListener: ((context: SelectionContext) => void) | null = null;

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;

  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });

  return { promise, resolve, reject };
}

beforeEach(() => {
  mockedBridge.ping.mockResolvedValue({
    host: 'browser-preview',
    version: 'dev',
  });
  mockedBridge.getSelectionContext.mockResolvedValue({
    hasSelection: true,
    workbookName: 'Quarterly Report.xlsx',
    sheetName: 'Sheet1',
    address: 'A1:C4',
    rowCount: 4,
    columnCount: 3,
    isContiguous: true,
    headerPreview: ['Name', 'Region', 'Amount'],
    sampleRows: [
      ['Project A', 'CN', '42'],
      ['Project B', 'US', '36'],
    ],
    warningMessage: null,
  });
  mockedBridge.getSessions.mockResolvedValue({
    activeSessionId: 'browser-preview-session',
    sessions: [
      {
        id: 'browser-preview-session',
        title: 'Browser preview',
        createdAtUtc: '2026-03-29T00:00:00.0000000Z',
        updatedAtUtc: '2026-03-29T00:00:00.0000000Z',
        messages: [],
      },
      {
        id: 'review-session',
        title: 'Review notes',
        createdAtUtc: '2026-03-29T01:00:00.0000000Z',
        updatedAtUtc: '2026-03-29T01:00:00.0000000Z',
        messages: [],
      },
    ],
  });
  mockedBridge.getSettings.mockResolvedValue({
    apiKey: '',
    baseUrl: 'https://api.example.com',
    model: 'gpt-5-mini',
  });
  mockedBridge.onSelectionContextChanged.mockImplementation((listener) => {
    selectionContextListener = listener;
    return () => {
      selectionContextListener = null;
    };
  });
  mockedBridge.saveSettings.mockImplementation(async (settings) => settings);
  mockedBridge.executeExcelCommand.mockResolvedValue({
    commandType: 'excel.readSelectionTable',
    requiresConfirmation: false,
    status: 'completed',
    message: 'Read selection from Sheet1 A1:C4.',
    table: {
      sheetName: 'Sheet1',
      address: 'A1:C4',
      headers: ['Name', 'Region', 'Amount'],
      rows: [
        ['Project A', 'CN', '42'],
        ['Project B', 'US', '36'],
      ],
    },
  });
  mockedBridge.runSkill.mockResolvedValue({
    route: 'chat',
    requiresConfirmation: false,
    status: 'completed',
    message: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
  });
  mockedBridge.runAgent.mockResolvedValue({
    route: 'chat',
    requiresConfirmation: false,
    status: 'completed',
    message: 'General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.',
  });
});

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
  selectionContextListener = null;
});

describe('App shell', () => {
  it('renders the expected task pane regions', async () => {
    const user = userEvent.setup();

    render(<App />);

    expect(
      screen.getByRole('complementary', { name: /session sidebar placeholder/i }),
    ).toBeInTheDocument();
    expect(screen.getByRole('banner', { name: /chat header/i })).toBeInTheDocument();
    expect(
      screen.getByRole('region', { name: /message thread/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/welcome to office agent/i),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('status', { name: /selection badge placeholder/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('textbox', { name: /message composer/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('button', { name: /settings/i }),
    ).toBeInTheDocument();
    expect(
      await screen.findByText(/connected to browser-preview \(dev\)/i),
    ).toBeInTheDocument();
    expect(
      await screen.findByText(/sheet1 · a1:c4/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/quarterly report\.xlsx/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/4 rows x 3 columns/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/contiguous selection/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/headers: name, region, amount/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/project a · cn · 42/i),
    ).toBeInTheDocument();
    expect(
      await screen.findByRole('button', { name: /browser preview/i }),
    ).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /settings/i }));

    expect(
      screen.getByRole('dialog', { name: /settings dialog/i }),
    ).toBeInTheDocument();
    expect(
      screen.getByRole('textbox', { name: /api key/i }),
    ).toHaveValue('');
    expect(
      screen.getByRole('textbox', { name: /base url/i }),
    ).toHaveValue('https://api.example.com');
    expect(
      screen.getByRole('textbox', { name: /model/i }),
    ).toHaveValue('gpt-5-mini');
  });

  it('switches the active session when a session chip is clicked', async () => {
    const user = userEvent.setup();

    render(<App />);
    const sidebar = screen.getByRole('complementary', { name: /session sidebar placeholder/i });

    expect(
      await screen.findByRole('heading', { name: /browser preview/i }),
    ).toBeInTheDocument();

    await user.click(within(sidebar).getByRole('button', { name: /review notes/i }));

    expect(
      screen.getByRole('heading', { name: /review notes/i }),
    ).toBeInTheDocument();
  });

  it('resets unsaved settings changes when the dialog is cancelled or closed', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    await user.clear(screen.getByRole('textbox', { name: /base url/i }));
    await user.type(screen.getByRole('textbox', { name: /base url/i }), 'https://changed.example.com');
    await user.click(screen.getByRole('button', { name: /cancel/i }));

    await user.click(screen.getByRole('button', { name: /settings/i }));
    expect(
      screen.getByRole('textbox', { name: /base url/i }),
    ).toHaveValue('https://api.example.com');

    await user.clear(screen.getByRole('textbox', { name: /base url/i }));
    await user.type(screen.getByRole('textbox', { name: /base url/i }), 'https://closed.example.com');
    await user.click(screen.getByRole('button', { name: /close/i }));

    await user.click(screen.getByRole('button', { name: /settings/i }));
    expect(
      screen.getByRole('textbox', { name: /base url/i }),
    ).toHaveValue('https://api.example.com');
  });

  it('shows an inline error when saving settings fails', async () => {
    const user = userEvent.setup();
    mockedBridge.saveSettings.mockRejectedValueOnce(new Error('Bridge write failed'));

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    await user.click(screen.getByRole('button', { name: /save/i }));

    expect(
      await screen.findByRole('alert'),
    ).toBeInTheDocument();
    expect(screen.getByRole('alert')).toHaveTextContent(/bridge write failed/i);
    expect(
      screen.getByRole('dialog', { name: /settings dialog/i }),
    ).toBeInTheDocument();
  });

  it('does not overwrite unsaved settings edits when settings load late', async () => {
    const user = userEvent.setup();
    const delayedSettings = createDeferred<{
      apiKey: string;
      baseUrl: string;
      model: string;
    }>();
    mockedBridge.getSettings.mockReturnValueOnce(delayedSettings.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    const baseUrlInput = screen.getByRole('textbox', { name: /base url/i });
    await user.clear(baseUrlInput);
    await user.type(baseUrlInput, 'https://draft.example.com');

    delayedSettings.resolve({
      apiKey: '',
      baseUrl: 'https://loaded.example.com',
      model: 'gpt-5-mini',
    });

    expect(
      await screen.findByDisplayValue('https://draft.example.com'),
    ).toBeInTheDocument();
  });

  it('treats the settings dialog as a modal and restores focus on close', async () => {
    const user = userEvent.setup();

    render(<App />);

    const settingsButton = screen.getByRole('button', { name: /settings/i });
    await user.click(settingsButton);

    const dialog = screen.getByRole('dialog', { name: /settings dialog/i });
    const apiKeyInput = screen.getByRole('textbox', { name: /api key/i });

    expect(dialog).toHaveAttribute('aria-modal', 'true');
    expect(apiKeyInput).toHaveFocus();

    await user.tab({ shift: true });
    expect(screen.getByRole('button', { name: /close/i })).toHaveFocus();

    await user.tab({ shift: true });
    expect(screen.getByRole('button', { name: /save/i })).toHaveFocus();

    await user.click(screen.getByRole('button', { name: /close/i }));
    expect(settingsButton).toHaveFocus();
  });

  it('disables save until settings finish loading', async () => {
    const user = userEvent.setup();
    const delayedSettings = createDeferred<{
      apiKey: string;
      baseUrl: string;
      model: string;
    }>();
    mockedBridge.getSettings.mockReturnValueOnce(delayedSettings.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    expect(screen.getByRole('button', { name: /save/i })).toBeDisabled();

    delayedSettings.resolve({
      apiKey: 'loaded-key',
      baseUrl: 'https://loaded.example.com',
      model: 'gpt-5-mini',
    });

    expect(
      await screen.findByDisplayValue('loaded-key'),
    ).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /save/i })).toBeEnabled();
  });

  it('keeps save disabled when settings fail to load', async () => {
    const user = userEvent.setup();
    mockedBridge.getSettings.mockRejectedValueOnce(new Error('Load failed'));

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));

    expect(screen.getByRole('button', { name: /save/i })).toBeDisabled();
    expect(await screen.findByRole('alert')).toHaveTextContent(/unable to load settings/i);
  });

  it('disables the settings form while save is in flight', async () => {
    const user = userEvent.setup();
    const pendingSave = createDeferred<{
      apiKey: string;
      baseUrl: string;
      model: string;
    }>();
    mockedBridge.saveSettings.mockReturnValueOnce(pendingSave.promise);

    render(<App />);

    await user.click(screen.getByRole('button', { name: /settings/i }));
    await user.click(screen.getByRole('button', { name: /save/i }));

    expect(screen.getByRole('textbox', { name: /api key/i })).toBeDisabled();
    expect(screen.getByRole('textbox', { name: /base url/i })).toBeDisabled();
    expect(screen.getByRole('textbox', { name: /model/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /close/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /cancel/i })).toBeDisabled();
    expect(screen.getByRole('button', { name: /save/i })).toBeDisabled();

    pendingSave.resolve({
      apiKey: '',
      baseUrl: 'https://api.example.com',
      model: 'gpt-5-mini',
    });

    expect(
      await screen.findByRole('button', { name: /settings/i }),
    ).toHaveFocus();
  });

  it('updates the selection badge when native selection events arrive', async () => {
    render(<App />);

    expect(
      await screen.findByText(/sheet1 · a1:c4/i),
    ).toBeInTheDocument();

    selectionContextListener?.({
      hasSelection: true,
      workbookName: 'Quarterly Report.xlsx',
      sheetName: 'Sheet2',
      address: 'B2:D5',
      rowCount: 4,
      columnCount: 3,
      isContiguous: false,
      headerPreview: [],
      sampleRows: [],
      warningMessage: 'Multiple selection areas are not supported yet.',
    });

    expect(
      await screen.findByText(/sheet2 · b2:d5/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/non-contiguous selection/i),
    ).toBeInTheDocument();
    expect(
      screen.getByText(/multiple selection areas are not supported yet/i),
    ).toBeInTheDocument();
  });

  it('shows the empty selection state when the native host has no selection context', async () => {
    mockedBridge.getSelectionContext.mockResolvedValueOnce({
      hasSelection: false,
      workbookName: '',
      sheetName: '',
      address: '',
      rowCount: 0,
      columnCount: 0,
      isContiguous: true,
      headerPreview: [],
      sampleRows: [],
      warningMessage: 'No selection available.',
    });

    render(<App />);

    expect(
      await screen.findByText(/no selection available/i),
    ).toBeInTheDocument();
  });

  it('submits read-selection commands without requiring confirmation', async () => {
    const user = userEvent.setup();

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /message composer/i }), 'read selection');
    await user.click(screen.getByRole('button', { name: /send/i }));

    expect(mockedBridge.executeExcelCommand).toHaveBeenCalledWith({
      commandType: 'excel.readSelectionTable',
      confirmed: false,
    });
    expect(screen.queryByText(/confirm excel action/i)).not.toBeInTheDocument();
    expect(
      await screen.findByText(/read selection from sheet1 a1:c4/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/name \| region \| amount/i)).toBeInTheDocument();
    expect(screen.getByText(/project a \| cn \| 42/i)).toBeInTheDocument();
  });

  it('routes plain natural language through the agent bridge', async () => {
    const user = userEvent.setup();
    mockedBridge.runAgent.mockResolvedValueOnce({
      route: 'chat',
      requiresConfirmation: false,
      status: 'completed',
      message: 'I can help with the current selection.',
    });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /message composer/i }), 'Create a summary sheet from the current selection');
    await user.click(screen.getByRole('button', { name: /send/i }));

    expect(mockedBridge.runAgent).toHaveBeenCalledWith({
      userInput: 'Create a summary sheet from the current selection',
      confirmed: false,
      sessionId: 'browser-preview-session',
    });
    expect(mockedBridge.runSkill).not.toHaveBeenCalled();
    expect(
      await screen.findByText(/i can help with the current selection/i),
    ).toBeInTheDocument();
  });

  it('renders a plan preview and confirms the frozen plan through the agent bridge', async () => {
    const user = userEvent.setup();
    const frozenPlan = {
      summary: 'Create a Summary sheet and write the selected rows.',
      steps: [
        {
          type: 'excel.addWorksheet',
          args: {
            newSheetName: 'Summary',
          },
        },
        {
          type: 'excel.writeRange',
          args: {
            targetAddress: 'Summary!A1:B3',
            values: [
              ['Name', 'Region'],
              ['Project A', 'CN'],
              ['Project B', 'US'],
            ],
          },
        },
      ],
    };

    mockedBridge.runAgent
      .mockResolvedValueOnce({
        route: 'plan',
        requiresConfirmation: true,
        status: 'preview',
        message: 'I prepared a plan. Review it before Excel is changed.',
        planner: {
          mode: 'plan',
          assistantMessage: 'I prepared a plan. Review it before Excel is changed.',
          plan: frozenPlan,
        },
      })
      .mockResolvedValueOnce({
        route: 'plan',
        requiresConfirmation: false,
        status: 'completed',
        message: 'Plan executed successfully.',
        journal: {
          hasFailures: false,
          errorMessage: '',
          steps: [
            {
              type: 'excel.addWorksheet',
              title: 'Add worksheet Summary',
              status: 'completed',
              message: 'Worksheet "Summary" created.',
              errorMessage: '',
            },
            {
              type: 'excel.writeRange',
              title: 'Write range Summary!A1:B3',
              status: 'completed',
              message: 'Wrote 3 row(s) to Summary!A1:B3.',
              errorMessage: '',
            },
          ],
        },
      });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /message composer/i }), 'Create a summary sheet from the current selection');
    await user.click(screen.getByRole('button', { name: /send/i }));

    expect(
      await screen.findByText(/create a summary sheet and write the selected rows/i),
    ).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /confirm/i }));

    expect(mockedBridge.runAgent).toHaveBeenNthCalledWith(1, {
      userInput: 'Create a summary sheet from the current selection',
      confirmed: false,
      sessionId: 'browser-preview-session',
    });
    expect(mockedBridge.runAgent).toHaveBeenNthCalledWith(2, {
      userInput: 'Create a summary sheet from the current selection',
      confirmed: true,
      sessionId: 'browser-preview-session',
      plan: frozenPlan,
    });
    expect(
      await screen.findByText(/plan executed successfully/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/completed · add worksheet summary/i)).toBeInTheDocument();
    expect(screen.getByText(/completed · write range summary!a1:b3/i)).toBeInTheDocument();
  });

  it('queues write commands for confirmation before executing them', async () => {
    const user = userEvent.setup();
    mockedBridge.executeExcelCommand
      .mockResolvedValueOnce({
        commandType: 'excel.addWorksheet',
        requiresConfirmation: true,
        status: 'preview',
        message: 'Confirm worksheet creation before Excel is modified.',
        preview: {
          title: 'Add worksheet',
          summary: 'Add worksheet "Summary"',
          details: ['Workbook: Quarterly Report.xlsx'],
        },
      })
      .mockResolvedValueOnce({
        commandType: 'excel.addWorksheet',
        requiresConfirmation: false,
        status: 'completed',
        message: 'Worksheet "Summary" created.',
      });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /message composer/i }), 'add sheet Summary');
    await user.click(screen.getByRole('button', { name: /send/i }));

    expect(
      await screen.findByText(/confirm excel action/i),
    ).toBeInTheDocument();
    expect(screen.getByText(/add worksheet "summary"/i)).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /confirm/i }));

    expect(mockedBridge.executeExcelCommand).toHaveBeenNthCalledWith(1, {
      commandType: 'excel.addWorksheet',
      newSheetName: 'Summary',
      confirmed: false,
    });
    expect(mockedBridge.executeExcelCommand).toHaveBeenNthCalledWith(2, {
      commandType: 'excel.addWorksheet',
      newSheetName: 'Summary',
      confirmed: true,
    });
    expect(
      await screen.findByText(/worksheet "summary" created/i),
    ).toBeInTheDocument();
  });

  it('disables the composer while a confirmation card is pending', async () => {
    const user = userEvent.setup();
    mockedBridge.runSkill.mockResolvedValueOnce({
      route: 'skill',
      skillName: 'upload_data',
      requiresConfirmation: true,
      status: 'preview',
      message: 'Review the upload payload before sending it to Project A.',
      preview: {
        title: 'Upload selected data',
        summary: 'Upload 2 row(s) to Project A',
        details: ['Source: Sheet1!A1:C3', 'Fields: Name, Region'],
      },
      uploadPreview: {
        projectName: 'Project A',
        sheetName: 'Sheet1',
        address: 'A1:C3',
        headers: ['Name', 'Region'],
        rows: [
          ['Project A', 'CN'],
          ['Project B', 'US'],
        ],
        records: [
          { Name: 'Project A', Region: 'CN' },
          { Name: 'Project B', Region: 'US' },
        ],
      },
    });

    render(<App />);

    const composer = screen.getByRole('textbox', { name: /message composer/i });
    const sendButton = screen.getByRole('button', { name: /send/i });
    await user.type(composer, '/upload_data upload selected data to Project A');
    await user.click(sendButton);

    expect(
      await screen.findByText(/confirm excel action/i),
    ).toBeInTheDocument();
    expect(composer).toBeDisabled();
    expect(sendButton).toBeDisabled();
    expect(mockedBridge.runSkill).toHaveBeenCalledTimes(1);
  });

  it('keeps command results in the session that launched them when the user switches threads mid-flight', async () => {
    const user = userEvent.setup();
    const pendingCommand = createDeferred<{
      commandType: string;
      requiresConfirmation: boolean;
      status: string;
      message: string;
    }>();
    mockedBridge.executeExcelCommand.mockReturnValueOnce(pendingCommand.promise);

    render(<App />);

    const sidebar = screen.getByRole('complementary', { name: /session sidebar placeholder/i });
    await screen.findByRole('heading', { name: /browser preview/i });

    await user.type(screen.getByRole('textbox', { name: /message composer/i }), 'read selection');
    await user.click(screen.getByRole('button', { name: /send/i }));
    await user.click(within(sidebar).getByRole('button', { name: /review notes/i }));

    expect(
      screen.getByRole('heading', { name: /review notes/i }),
    ).toBeInTheDocument();

    pendingCommand.resolve({
      commandType: 'excel.readSelectionTable',
      requiresConfirmation: false,
      status: 'completed',
      message: 'Read selection from Sheet1 A1:C4.',
    });

    expect(
      screen.queryByText(/read selection from sheet1 a1:c4/i),
    ).not.toBeInTheDocument();

    await user.click(within(sidebar).getByRole('button', { name: /browser preview/i }));

    expect(
      await screen.findByText(/read selection from sheet1 a1:c4/i),
    ).toBeInTheDocument();
  });

  it('routes explicit slash upload_data through the skill bridge and confirms with the returned preview payload', async () => {
    const user = userEvent.setup();
    const uploadPreview = {
      projectName: '项目A',
      sheetName: 'Sheet1',
      address: 'A1:C3',
      headers: ['Name', 'Region'],
      rows: [
        ['Project A', 'CN'],
        ['Project B', 'US'],
      ],
      records: [
        { Name: 'Project A', Region: 'CN' },
        { Name: 'Project B', Region: 'US' },
      ],
    };

    mockedBridge.runSkill
      .mockResolvedValueOnce({
        route: 'skill',
        skillName: 'upload_data',
        requiresConfirmation: true,
        status: 'preview',
        message: 'Review the upload payload before sending it to 项目A.',
        preview: {
          title: 'Upload selected data',
          summary: 'Upload 2 row(s) to 项目A',
          details: ['Source: Sheet1!A1:C3', 'Fields: Name, Region'],
        },
        uploadPreview,
      })
      .mockResolvedValueOnce({
        route: 'skill',
        skillName: 'upload_data',
        requiresConfirmation: false,
        status: 'completed',
        message: 'Uploaded 2 row(s) to 项目A.',
        uploadPreview,
      });

    render(<App />);

    await user.type(screen.getByRole('textbox', { name: /message composer/i }), '把选中数据上传到项目A');
    await user.click(screen.getByRole('button', { name: /send/i }));

    expect(mockedBridge.runSkill).toHaveBeenNthCalledWith(1, {
      userInput: '把选中数据上传到项目A',
      confirmed: false,
    });
    expect(
      await screen.findByText(/upload 2 row\(s\) to 项目a/i),
    ).toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: /confirm/i }));

    expect(mockedBridge.runSkill).toHaveBeenNthCalledWith(2, {
      userInput: '把选中数据上传到项目A',
      skillName: 'upload_data',
      confirmed: true,
      uploadPreview,
    });
    expect(
      await screen.findByText(/uploaded 2 row\(s\) to 项目a/i),
    ).toBeInTheDocument();
  });
});
