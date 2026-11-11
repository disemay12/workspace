/**
 * Страница: появление блоков, шапка по прокрутке, липкий кадр по шагам,
 * отсчёт измеренных величин, фон за курсором и просмотр скриншотов.
 * Без JS весь контент виден, первый кадр секции остаётся на месте,
 * а скриншоты открываются обычной ссылкой на файл.
 */
(() => {
  const root = document.documentElement;
  const calm = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const all = (selector) => [...document.querySelectorAll(selector)];

  if (!calm) root.classList.add('js');

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

  // шапка
  const nav = document.getElementById('nav');
  if (nav) {
    const sync = () => nav.classList.toggle('is-scrolled', window.scrollY > 6);
    sync();
    window.addEventListener('scroll', sync, { passive: true });
  }

  // текущий раздел в шапке
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

  // липкий кадр по шагам
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

  // фон, который ходит за курсором
  const bg = document.querySelector('.bg');
  const parallax = document.querySelector('.hero__parallax');
  const finePointer = window.matchMedia('(hover: hover) and (pointer: fine)').matches;

  if (bg && finePointer && !calm) {
    let targetX = window.innerWidth / 2;
    let targetY = window.innerHeight * 0.32;
    let warmX = targetX;
    let warmY = targetY;
    let coolX = targetX;
    let coolY = targetY;
    let shiftX = 0;
    let shiftY = 0;

    window.addEventListener(
      'pointermove',
      (event) => {
        if (event.pointerType !== 'mouse') return;
        targetX = event.clientX;
        targetY = event.clientY;
        bg.classList.add('is-live');
      },
      { passive: true }
    );

    const frame = () => {
      warmX += (targetX - warmX) * 0.24;
      warmY += (targetY - warmY) * 0.24;
      coolX += (targetX - coolX) * 0.07;
      coolY += (targetY - coolY) * 0.07;
      shiftX += ((targetX / window.innerWidth - 0.5) * 18 - shiftX) * 0.1;
      shiftY += ((targetY / window.innerHeight - 0.5) * 14 - shiftY) * 0.1;

      bg.style.setProperty('--mx', `${warmX.toFixed(1)}px`);
      bg.style.setProperty('--my', `${warmY.toFixed(1)}px`);
      bg.style.setProperty('--lx', `${coolX.toFixed(1)}px`);
      bg.style.setProperty('--ly', `${coolY.toFixed(1)}px`);

      if (parallax) {
        parallax.style.setProperty('--hx', `${shiftX.toFixed(2)}px`);
        parallax.style.setProperty('--hy', `${shiftY.toFixed(2)}px`);
      }

      requestAnimationFrame(frame);
    };

    requestAnimationFrame(frame);
  }

  // просмотр скриншотов
  const viewer = document.getElementById('viewer');
  const gallery = all('[data-gallery]');

  if (viewer && gallery.length && typeof viewer.showModal === 'function') {
    const vImage = viewer.querySelector('.viewer__img');
    const vCaption = viewer.querySelector('.viewer__cap');
    const shotOf = (figure) => figure.querySelector('.pin__img.is-active') || figure.querySelector('img');
    const captionOf = (figure) => {
      const caption = figure.querySelector('figcaption');
      return (caption ? caption.textContent : shotOf(figure).alt).trim();
    };
    let current = 0;

    const source = (figure) => {
      const image = shotOf(figure);
      return image.currentSrc || image.src;
    };

    const paint = (index) => {
      current = (index + gallery.length) % gallery.length;
      const image = shotOf(gallery[current]);
      vImage.src = source(gallery[current]);
      vImage.alt = image.alt;
      vCaption.textContent = captionOf(gallery[current]);
    };

    // один и тот же кадр может стоять дважды (первый экран и шаг секции)
    const shift = (direction) => {
      for (let step = 1; step <= gallery.length; step += 1) {
        const index = (((current + direction * step) % gallery.length) + gallery.length) % gallery.length;
        if (source(gallery[index]) === vImage.src) continue;
        paint(index);
        return;
      }
    };

    const grow = (origin) => {
      if (calm || !origin || !vImage.animate) return;
      const from = origin.getBoundingClientRect();
      const to = vImage.getBoundingClientRect();
      if (!from.width || !to.width) return;
      const scaleX = from.width / to.width;
      const scaleY = from.height / to.height;
      const dx = from.left + from.width / 2 - (to.left + to.width / 2);
      const dy = from.top + from.height / 2 - (to.top + to.height / 2);
      vImage.animate(
        [
          { transform: `translate(${dx}px, ${dy}px) scale(${scaleX}, ${scaleY})` },
          { transform: 'none' },
        ],
        { duration: 460, easing: 'cubic-bezier(0.16, 1, 0.3, 1)' }
      );
    };

    const open = (index, origin) => {
      paint(index);
      viewer.showModal();
      root.classList.add('is-locked');
      grow(origin);
      const next = viewer.querySelector('[data-dir="1"]');
      if (next) next.focus({ preventScroll: true });
    };

    gallery.forEach((figure, index) => {
      const trigger = figure.querySelector('.shot__open');
      if (!trigger) return;
      trigger.addEventListener('click', () => open(index, trigger));
    });

    viewer.addEventListener('click', (event) => {
      if (event.target === viewer) viewer.close();
    });

    for (const button of all('.viewer__btn')) {
      button.addEventListener('click', () => {
        if (button.hasAttribute('data-close')) {
          viewer.close();
          return;
        }
        shift(Number(button.dataset.dir));
      });
    }

    viewer.addEventListener('keydown', (event) => {
      if (event.key === 'ArrowRight') {
        shift(1);
        event.preventDefault();
      } else if (event.key === 'ArrowLeft') {
        shift(-1);
        event.preventDefault();
      }
    });

    viewer.addEventListener('close', () => root.classList.remove('is-locked'));
  }
  // форма предложений: проверка полей и отправка на почту через FormSubmit
  const form = document.getElementById('feedback-form');

  if (form) {
    // адрес получателя не держим в разметке открытым текстом
    const address = String.fromCharCode(
      103, 101, 116, 117, 112, 51, 50, 50, 64, 103, 109, 97, 105, 108, 46, 99, 111, 109
    );
    const endpoint = `https://formsubmit.co/ajax/${address}`;
    form.setAttribute('action', `https://formsubmit.co/${address}`);

    const status = document.getElementById('f-status');
    const submit = form.querySelector('button[type="submit"]');
    const idea = form.elements.idea;
    const counter = document.getElementById('f-count');
    const readyAt = Date.now();
    const cooldown = 90_000;
    const storageKey = 'workspace-feedback-sent';

    const setStatus = (text, kind) => {
      status.textContent = text;
      status.className = `form__status${kind ? ` is-${kind}` : ''}`;
    };

    const setError = (name, text) => {
      const box = form.querySelector(`[data-error-for="${name}"]`);
      if (box) box.textContent = text;
      if (form.elements[name]) form.elements[name].setAttribute('aria-invalid', text ? 'true' : 'false');
    };

    const rules = {
      name: (value) => (value.trim().length >= 2 && value.trim().length <= 60 ? '' : 'Напишите имя: от 2 до 60 символов.'),
      email: (value) =>
        /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(value.trim()) ? '' : 'Проверьте адрес: нужен вид name@example.com.',
      idea: (value) =>
        value.trim().length >= 10 && value.trim().length <= 1200 ? '' : 'Опишите предложение: от 10 до 1200 символов.',
    };

    // одноразовые почты: такие адреса не принимаем
    const disposable = [
      '10minutemail.com', '20minutemail.com', '33mail.com', 'anonbox.net', 'burnermail.io',
      'discard.email', 'dispostable.com', 'emailondeck.com', 'fakeinbox.com', 'getnada.com',
      'guerrillamail.com', 'guerrillamail.net', 'inboxkitten.com', 'mail.tm', 'mail7.io',
      'mailcatch.com', 'maildrop.cc', 'mailinator.com', 'mailnesia.com', 'mintemail.com',
      'mohmal.com', 'mytemp.email', 'sharklasers.com', 'spam4.me', 'spamgourmet.com',
      'temp-mail.io', 'temp-mail.org', 'tempail.com', 'tempmail.com', 'tempmailo.com',
      'tempr.email', 'throwawaymail.com', 'trashmail.com', 'yopmail.com', 'yopmail.fr',
    ];

    // домен должен уметь принимать почту: спрашиваем MX через DNS-over-HTTPS
    const checkDomain = async (email) => {
      const domain = email.split('@')[1].toLowerCase();
      if (disposable.includes(domain)) return 'Одноразовая почта не подойдёт: укажите постоянный адрес.';

      const ask = async (url) => {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), 4000);
        try {
          const response = await fetch(url, { headers: { Accept: 'application/dns-json' }, signal: controller.signal });
          if (!response.ok) return null;
          const data = await response.json();
          return (data.Answer || []).some((record) => record.type === 15);
        } catch {
          return null;
        } finally {
          clearTimeout(timer);
        }
      };

      const hasMail =
        (await ask(`https://cloudflare-dns.com/dns-query?name=${encodeURIComponent(domain)}&type=MX`)) ??
        (await ask(`https://dns.google/resolve?name=${encodeURIComponent(domain)}&type=MX`));

      if (hasMail === false) return 'У этого домена нет почтового сервера: письмо не дойдёт. Проверьте адрес.';
      return '';
    };

    const readSent = () => {
      try {
        return JSON.parse(localStorage.getItem('workspace-feedback-mails') || '{}') || {};
      } catch {
        return {};
      }
    };

    const alreadySent = (email) => {
      const map = readSent();
      const key = email.trim().toLowerCase();
      return Object.prototype.hasOwnProperty.call(map, key) && Date.now() - map[key] < 30 * 24 * 3600 * 1000;
    };

    const rememberSent = (email) => {
      const map = readSent();
      map[email.trim().toLowerCase()] = Date.now();
      try {
        localStorage.setItem('workspace-feedback-mails', JSON.stringify(map));
      } catch {
        // приватный режим: отметку сохранить не удалось
      }
    };

    if (idea && counter) {
      counter.textContent = String(idea.value.length);
      idea.addEventListener('input', () => {
        counter.textContent = String(idea.value.length);
      });
    }

    form.addEventListener('submit', async (event) => {
      event.preventDefault();

      if (form.elements._honey && form.elements._honey.value) {
        setStatus('Проверка не пройдена: похоже на автоматическую отправку.', 'error');
        return;
      }

      if (Date.now() - readyAt < 2500) {
        setStatus('Слишком быстро. Проверьте текст и отправьте ещё раз.', 'error');
        return;
      }

      let firstInvalid = null;
      for (const [name, rule] of Object.entries(rules)) {
        const message = rule(form.elements[name].value);
        setError(name, message);
        if (message && !firstInvalid) firstInvalid = form.elements[name];
      }

      if (firstInvalid) {
        firstInvalid.focus();
        setStatus('Проверьте поля формы.', 'error');
        return;
      }

      const emailValue = form.elements.email.value.trim();

      if (alreadySent(emailValue)) {
        setError('email', 'С этого адреса уже отправляли: одно обращение на почту.');
        form.elements.email.focus();
        setStatus('Одно обращение на адрес.', 'error');
        return;
      }

      setStatus('Проверяем адрес…');
      const domainProblem = await checkDomain(emailValue);
      if (domainProblem) {
        setError('email', domainProblem);
        form.elements.email.focus();
        setStatus('Проверьте адрес почты.', 'error');
        return;
      }

      let last = 0;
      try {
        last = Number(localStorage.getItem(storageKey) || 0);
      } catch {
        last = 0;
      }
      if (Date.now() - last < cooldown) {
        setStatus('Обращение уже отправлено. Подождите полторы минуты.', 'error');
        return;
      }

      submit.disabled = true;
      setStatus('Отправляем…');
      if (form.elements._replyto) form.elements._replyto.value = emailValue;

      try {
        const response = await fetch(endpoint, {
          method: 'POST',
          headers: { Accept: 'application/json' },
          body: new FormData(form),
        });
        const data = await response.json().catch(() => ({}));
        const ok = response.ok && String(data.success) !== 'false';
        if (!ok) throw new Error(data.message || `HTTP ${response.status}`);

        try {
          localStorage.setItem(storageKey, String(Date.now()));
        } catch {
          // приватный режим: ограничение частоты не сохранится
        }
        rememberSent(emailValue);

        if (form.elements._replyto) form.elements._replyto.value = '';

        form.reset();
        if (counter) counter.textContent = '0';
        for (const name of Object.keys(rules)) setError(name, '');
        setStatus('Отправлено. Спасибо, ответим на указанную почту.', 'ok');
      } catch (error) {
        const reason = String((error && error.message) || error);
        const text = /activation/i.test(reason)
          ? 'Отправка пока не подтверждена владельцем формы. Попробуйте позже.'
          : `Не получилось отправить: ${reason.slice(0, 140)}. Попробуйте ещё раз.`;
        setStatus(text, 'error');
      } finally {
        submit.disabled = false;
      }
    });
  }
})();
