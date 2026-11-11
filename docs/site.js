/**
 * Движение страницы: класс .js открывает появление блоков, наблюдатели
 * включают их, шапка реагирует на прокрутку, липкий кадр меняется по шагам,
 * измеренные величины один раз досчитываются до настоящего значения.
 * Без JS весь контент виден сразу, в липкой секции остаётся первый кадр.
 */
(() => {
  const root = document.documentElement;
  const calm = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  if (!calm) root.classList.add('js');

  const all = (selector) => [...document.querySelectorAll(selector)];

  // появление блоков
  if (calm || !('IntersectionObserver' in window)) {
    for (const el of all('.reveal')) el.classList.add('is-in');
  } else {
    const reveal = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          entry.target.classList.add('is-in');
          reveal.unobserve(entry.target);
        }
      },
      { rootMargin: '0px 0px -6% 0px', threshold: 0.12 }
    );
    for (const el of all('.reveal')) reveal.observe(el);
  }

  // шапка: уплотняется после прокрутки
  const nav = document.getElementById('nav');
  if (nav) {
    const sync = () => nav.classList.toggle('is-scrolled', window.scrollY > 6);
    sync();
    window.addEventListener('scroll', sync, { passive: true });
  }

  // текущий раздел подсвечивается в шапке
  const links = all('.nav__links a[href^="#"]');
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

  // липкий кадр в секции редакторов меняется по шагам
  const steps = all('.step[data-shot]');
  const pins = all('.pin__img[data-shot-img]');
  const activate = (name) => {
    for (const pin of pins) pin.classList.toggle('is-active', pin.dataset.shotImg === name);
    for (const step of steps) step.classList.toggle('is-active', step.dataset.shot === name);
  };

  if (steps.length && pins.length && 'IntersectionObserver' in window) {
    const watcher = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) activate(entry.target.dataset.shot);
        }
      },
      { rootMargin: '-42% 0px -42% 0px', threshold: 0 }
    );
    steps.forEach((step) => watcher.observe(step));
  }

  // измеренные величины
  const counters = all('[data-count]');
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
