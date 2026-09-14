// Shared chrome behavior for the Learning Path section:
// scroll-progress bar under the nav, and the mobile nav toggle.
// Include on every learning-path page (index + each lesson).
(function () {
  const root = document.documentElement;
  function updateScroll() {
    const max = document.documentElement.scrollHeight - window.innerHeight;
    const p = max > 0 ? window.scrollY / max : 0;
    root.style.setProperty('--scroll-p', p.toFixed(4));
  }
  window.addEventListener('scroll', updateScroll, { passive: true });
  updateScroll();

  const navToggle = document.querySelector('.nav-toggle');
  const navLinks = document.getElementById('nav-links');
  if (navToggle && navLinks) {
    navToggle.addEventListener('click', () => {
      const open = navLinks.classList.toggle('open');
      navToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    navLinks.querySelectorAll('a').forEach((a) => {
      a.addEventListener('click', () => {
        navLinks.classList.remove('open');
        navToggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // Share bar: wire up LinkedIn/X/Facebook/WhatsApp share links and the
  // copy-link button from a single `data-share-url` / `data-share-text`
  // pair on the .share-bar wrapper, so a lesson page only has to declare
  // those two attributes once.
  document.querySelectorAll('[data-share]').forEach((bar) => {
    const url = bar.dataset.shareUrl || window.location.href;
    const text = bar.dataset.shareText || document.title;
    const encodedUrl = encodeURIComponent(url);
    const encodedText = encodeURIComponent(text);

    const targets = {
      linkedin: `https://www.linkedin.com/sharing/share-offsite/?url=${encodedUrl}`,
      x: `https://twitter.com/intent/tweet?url=${encodedUrl}&text=${encodedText}`,
      facebook: `https://www.facebook.com/sharer/sharer.php?u=${encodedUrl}`,
      whatsapp: `https://api.whatsapp.com/send?text=${encodedText}%20${encodedUrl}`,
    };
    Object.keys(targets).forEach((key) => {
      const el = bar.querySelector(`[data-share-${key}]`);
      if (el) el.href = targets[key];
    });

    const copyBtn = bar.querySelector('[data-copy-link]');
    if (copyBtn) {
      copyBtn.addEventListener('click', async () => {
        try {
          await navigator.clipboard.writeText(url);
        } catch (err) {
          // Fallback for browsers without clipboard API access.
          const tmp = document.createElement('textarea');
          tmp.value = url;
          tmp.style.position = 'fixed';
          tmp.style.opacity = '0';
          document.body.appendChild(tmp);
          tmp.select();
          document.execCommand('copy');
          document.body.removeChild(tmp);
        }
        copyBtn.classList.add('copied');
        clearTimeout(copyBtn._resetTimer);
        copyBtn._resetTimer = setTimeout(() => copyBtn.classList.remove('copied'), 2000);
      });
    }
  });
})();
