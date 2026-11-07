/*
 * ============================================================
 *  CAPTCHA – code + word, then confirm
 * ============================================================
 * Used in Step 1 (register.html) and Step 4 (application.html).
 *
 * 1. The visitor creates a 5-character alphanumeric code (letters/numbers),
 *    e.g. "ab12X".
 * 2. The visitor writes a word of their choice, any length, letters only,
 *    e.g. "beach".
 * 3. The visitor types both together (code + word), e.g. "ab12Xbeach".
 *    check() verifies that this confirmation matches EXACTLY the code
 *    followed by the word (upper/lower case matter).
 * ============================================================
 */
var Captcha = (function () {
  var CODE_RE = /^[A-Za-z0-9]{5}$/;          // exactly 5 letters/numbers
  var WORD_RE = /^[A-Za-zÀ-ÿ]+$/;            // letters only, any length

  function mount(container) {
    container.innerHTML =
      '<div class="captcha-campo">' +
        '<label class="etiqueta">1. Create a 5-character code (letters and numbers) <span class="req">*</span></label>' +
        '<input type="text" class="c-codigo" maxlength="5" autocomplete="off" placeholder="e.g. ab12X">' +
        '<span class="error"></span>' +
      '</div>' +
      '<div class="captcha-campo">' +
        '<label class="etiqueta">2. Choose a word (letters only, any length) <span class="req">*</span></label>' +
        '<input type="text" class="c-palabra" autocomplete="off" placeholder="e.g. beach">' +
        '<span class="error"></span>' +
      '</div>' +
      '<div class="captcha-campo">' +
        '<label class="etiqueta">3. Type your code followed by your word, no spaces <span class="req">*</span></label>' +
        '<input type="text" class="c-confirmar" autocomplete="off" placeholder="e.g. ab12Xbeach">' +
        '<span class="error"></span>' +
      '</div>';

    var code = container.querySelector('.c-codigo');
    var word = container.querySelector('.c-palabra');
    var confirm = container.querySelector('.c-confirmar');

    function setError(input, msg) {
      input.parentNode.querySelector('.error').textContent = msg;
      input.classList.toggle('invalido', !!msg);
      return !msg;
    }

    // Validates the three boxes; the third must equal code + word.
    function check() {
      var ok1 = setError(code, !code.value ? 'Create your code.'
        : !CODE_RE.test(code.value) ? 'The code must have exactly 5 letters or numbers.' : '');
      var ok2 = setError(word, !word.value ? 'Write your user\'s name.'
        : !WORD_RE.test(word.value) ? 'Use letters only (no numbers, spaces or symbols).' : '');
      var ok3 = setError(confirm, !confirm.value ? 'Type your code and your word together.'
        : (ok1 && ok2 && confirm.value !== code.value + word.value) ? 'It doesn\'t match your code followed by your word.' : '');
      var first = !ok1 ? code : !ok2 ? word : !ok3 ? confirm : null;
      if (first) { first.focus(); return false; }
      return true;
    }

    // Clear an error as soon as the visitor edits that box.
    [code, word, confirm].forEach(function (i) {
      i.addEventListener('input', function () { setError(i, ''); });
    });

    return { check: check };
  }

  return { mount: mount };
})();
