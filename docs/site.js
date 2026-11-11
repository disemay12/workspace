/**
 * Движение страницы: класс .js открывает появление секций, наблюдатель
 * включает их, шапка реагирует на прокрутку, измеренные величины один раз
 * досчитываются до настоящего значения. Без JS весь контент виден сразу.
 */
(() => {
  const root = document.documentElement;
  const calm = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  if (!calm) root.classList.add('js');

  const revealAll = () => {
    for (const el of document.querySelectorAll('.reveal')) el.classList.add('is-in');
  };

  if (calm || !('IntersectionObserver' in window)) {
    revealAll();
  } else {
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          entry.target.classList.add('is-in');
          observer.unobserve(entry.target);
        }
      },
      { rootMargin: '0px 0px -6% 0px', threshold: 0.12 }
    );
    for (const el of document.querySelectorAll('.reveal')) observer.observe(el);
  }

  const nav = document.getElementById('nav');
  if (nav) {
    const sync = () => nav.classList.toggle('is-scrolled', window.scrollY > 6);
    sync();
    window.addEventListener('scroll', sync, { passive: true });
  }

  // текущий раздел подсвечивается в шапке
  const links = [...document.querySelectorAll('.nav__links a[href^="#"]')];
  const sections = links
    .map((link) => document.getElementById(link.getAttribute('href').slice(1)))
    .filter(Boolean);

  if (sections.length && 'IntersectionObserver' in window) {
    const inBand = new Set();
    const spy = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) inBand.add(entry.target.id);
          else inBand.delete(entry.target.id);
        }
        const current = sections.find((section) => inBand.has(section.id));
        for (const link of links) {
          link.classList.toggle('is-current', Boolean(current) && link.getAttribute('href') === `#${current.id}`);
        }
      },
      { rootMargin: '-45% 0px -45% 0px', threshold: 0 }
    );
    sections.forEach((section) => spy.observe(section));
  }

  const counters = [...document.querySelectorAll('[data-count]')];
  const settle = (el) => {
    el.textContent = el.dataset.count + (el.dataset.suffix || '');
  };
  const count = (el) => {
    const target = Number(el.dataset.count);
    const suffix = el.dataset.suffix || '';
    const start = performance.now();
    const step = (now) => {
      const progress = Math.min(1, (now - start) / 900);
      el.textContent = Math.round(target * (1 - Math.pow(1 - progress, 3))) + suffix;
      if (progress < 1) requestAnimationFrame(step);
      else settle(el);
    };
    requestAnimationFrame(step);
  };

  if (calm || !('IntersectionObserver' in window)) {
    counters.forEach(settle);
  } else {
    const watcher = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          count(entry.target);
          watcher.unobserve(entry.target);
        }
      },
      { threshold: 0.6 }
    );
    counters.forEach((el) => watcher.observe(el));
  }
})();
