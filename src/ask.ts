/**
 * One question, one answer.
 *
 * `window.prompt` would do the job and would be the only piece of chrome in
 * the program that looks like it belongs to the browser rather than to the
 * document. This is the command palette's shape with a single field, and it
 * keeps the same promise: the caret goes back where it was before anything
 * runs, so an answer that applies to a selection still has one.
 */

let host: HTMLElement | null = null;
let savedRange: Range | null = null;

export function isAskOpen(): boolean {
  return !!host;
}

function close(restore: boolean): void {
  host?.remove();
  host = null;
  if (restore && savedRange) {
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(savedRange);
  }
  savedRange = null;
}

export interface AskOptions {
  title: string;
  placeholder?: string;
  value?: string;
  confirm?: string;
  /** Returns an error message to show, or null when the value is good. */
  validate?: (value: string) => string | null;
}

/**
 * Resolves with the answer, or null when dismissed.
 *
 * The selection is captured before the field takes focus and restored before
 * the promise resolves, so the caller can act on it immediately.
 */
export function ask(opts: AskOptions): Promise<string | null> {
  close(false);
  const sel = window.getSelection();
  savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0).cloneRange() : null;

  return new Promise((resolve) => {
    const done = (value: string | null): void => {
      close(true);
      resolve(value);
    };

    host = document.createElement('div');
    host.className = 'cmd-scrim';
    host.addEventListener('mousedown', (e) => {
      if (e.target === host) {
        e.preventDefault();
        done(null);
      }
    });

    const panel = document.createElement('div');
    panel.className = 'cmd-panel ask-panel';

    const head = document.createElement('div');
    head.className = 'cmd-head';
    const mark = document.createElement('span');
    mark.className = 'cmd-mark';
    const label = document.createElement('span');
    label.className = 'ask-title';
    label.textContent = opts.title;
    head.append(mark, label);
    panel.appendChild(head);

    const body = document.createElement('div');
    body.className = 'ask-body';
    const input = document.createElement('input');
    input.className = 'cmd-input ask-input';
    input.type = 'text';
    input.placeholder = opts.placeholder ?? '';
    input.value = opts.value ?? '';
    input.spellcheck = false;
    input.autocomplete = 'off';
    body.appendChild(input);

    const error = document.createElement('div');
    error.className = 'ask-error';
    body.appendChild(error);
    panel.appendChild(body);

    const foot = document.createElement('div');
    foot.className = 'ask-foot';
    const hint = document.createElement('span');
    hint.className = 'ask-hint';
    hint.textContent = 'ENTER TO CONFIRM · ESC TO CANCEL';
    foot.appendChild(hint);
    panel.appendChild(foot);

    host.appendChild(panel);
    document.body.appendChild(host);

    const submit = (): void => {
      const value = input.value.trim();
      const bad = opts.validate ? opts.validate(value) : null;
      if (bad) {
        error.textContent = bad;
        input.classList.add('bad');
        return;
      }
      done(value);
    };

    input.addEventListener('input', () => {
      error.textContent = '';
      input.classList.remove('bad');
    });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        done(null);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        submit();
      }
    });

    input.focus();
    input.select();
  });
}
