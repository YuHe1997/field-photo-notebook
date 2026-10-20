// Form ergonomics for phones: auto-growing text boxes and keyboard awareness.
// Presentation only; never changes what a field contains.

const NON_TEXT_INPUTS = new Set(['checkbox', 'radio', 'button', 'submit', 'reset', 'file', 'range', 'color', 'image', 'hidden']);

// Resize a textarea to show all of its text. Skips hidden elements (height would measure 0).
export function fitTextarea(el) {
  if (!el.getClientRects().length) return;
  el.style.height = 'auto';
  const border = el.offsetHeight - el.clientHeight;
  el.style.height = `${el.scrollHeight + border}px`;
}

// While the user types, keep the box's bottom edge where it is on screen so growth
// extends upward and the caret never ends up behind the phone keyboard.
function fitWhileTyping(el) {
  if (document.activeElement !== el) return fitTextarea(el);
  const before = el.getBoundingClientRect().bottom;
  fitTextarea(el);
  const shift = el.getBoundingClientRect().bottom - before;
  if (shift) window.scrollBy(0, shift);
}

export function initAutoGrow(textareas) {
  const valueProperty = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value');
  const widths = new WeakMap();
  const observer = typeof ResizeObserver === 'function'
    ? new ResizeObserver(entries => {
      for (const { target } of entries) {
        const width = target.clientWidth;
        if (widths.get(target) === width) continue; // our own height change; nothing to do
        widths.set(target, width);
        fitTextarea(target);
      }
    })
    : null;
  for (const el of textareas) {
    if (!el || el.dataset.autoGrow) continue;
    el.dataset.autoGrow = 'true';
    el.classList.add('auto-grow');
    // Values are also set from code (saved drafts, templates, numbering); refit those too.
    Object.defineProperty(el, 'value', {
      configurable: true,
      get() { return valueProperty.get.call(this); },
      set(next) { valueProperty.set.call(this, next); fitTextarea(this); },
    });
    el.addEventListener('input', () => fitWhileTyping(el));
    observer?.observe(el);
    fitTextarea(el);
  }
}

// A fixed bottom bar jumps up above the keyboard on some phones. Flag keyboard-open so
// the stylesheet can tuck the bar away while the user is typing.
export function initKeyboardState() {
  const viewport = window.visualViewport;
  let baseline = window.innerHeight;
  const editing = () => {
    const el = document.activeElement;
    return !!el && (el.tagName === 'TEXTAREA' || (el.tagName === 'INPUT' && !NON_TEXT_INPUTS.has(el.type)));
  };
  const update = () => {
    const visible = viewport ? viewport.height * (viewport.scale || 1) : window.innerHeight;
    const isEditing = editing();
    if (!isEditing) baseline = Math.max(window.innerHeight, visible);
    document.body.classList.toggle('keyboard-open', isEditing && visible < baseline - 120);
  };
  viewport?.addEventListener('resize', update);
  window.addEventListener('resize', update);
  document.addEventListener('focusin', update);
  document.addEventListener('focusout', () => requestAnimationFrame(update));
  update();
}

export function initFormFit() {
  initAutoGrow([...document.querySelectorAll('#monForm textarea'), document.getElementById('engineering')]);
  initKeyboardState();
}
