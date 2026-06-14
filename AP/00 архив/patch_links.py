import sys, re
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

index = Path(__file__).parent / "netlify_deploy" / "index.html"

BLOCKER = """
<script>
(function () {
  // Block window.open (popup links)
  window.open = function () { return null; };

  // Block all form submissions
  document.addEventListener('submit', function (e) { e.preventDefault(); }, true);

  // Neutralise every <a> found now and after Vue hydration
  function neutralise() {
    document.querySelectorAll('a[href]').forEach(function (a) {
      var h = a.getAttribute('href');
      if (h && h !== '#' && !h.startsWith('javascript:') && !h.startsWith('mailto:')) {
        a.setAttribute('href', 'javascript:void(0)');
        a.style.cursor = 'default';
      }
    });
  }

  // Intercept clicks before they bubble (catches dynamically rendered Vue links)
  document.addEventListener('click', function (e) {
    var el = e.target;
    while (el) {
      if (el.tagName === 'A' || el.tagName === 'BUTTON') {
        var href = el.getAttribute && el.getAttribute('href');
        // allow only in-page anchors
        if (!href || (href !== '#' && !href.startsWith('#') && !href.startsWith('javascript:'))) {
          e.preventDefault();
          e.stopImmediatePropagation();
        }
        break;
      }
      el = el.parentElement;
    }
  }, true);

  // Run now + after DOMContentLoaded + watch for Vue-injected nodes
  neutralise();
  document.addEventListener('DOMContentLoaded', neutralise);
  var obs = new MutationObserver(neutralise);
  function startObs() {
    obs.observe(document.body || document.documentElement, { childList: true, subtree: true });
  }
  if (document.body) { startObs(); }
  else { document.addEventListener('DOMContentLoaded', startObs); }
})();
</script>
"""

html = index.read_text(encoding='utf-8')

if 'neutralise' in html:
    print("Патч уже применён.")
else:
    html = html.replace('</body>', BLOCKER + '\n</body>', 1)
    if '</body>' not in html:
        html += BLOCKER
    index.write_text(html, encoding='utf-8')
    print(f"Готово: {index}")
    print("Все ссылки/кнопки заблокированы — страница никуда не ведёт.")
