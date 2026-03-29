import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import App from './App';

describe('App shell', () => {
  it('renders the expected task pane regions', async () => {
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
  });
});
