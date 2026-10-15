/**
 * Проверка логики подтверждения выхода без запуска окна Electron.
 *
 * Повторяет конечный автомат из electron/main.ts: перехват закрытия,
 * отправка запроса в рендерер, ответ пользователя и повторная попытка.
 * Именно эта логика защищает несохранённые файлы от потери.
 */

/** Состояние перехвата закрытия — как в main.ts. */
interface CloseState {
  allowClose: boolean;
  closeRequestPending: boolean;
}

function newState(): CloseState {
  return { allowClose: false, closeRequestPending: false };
}

/** Обработчик события 'close' окна. Возвращает, было ли закрытие отменено. */
function onWindowClose(
  state: CloseState,
  rendererAlive: boolean
): { prevented: boolean; sentRequest: boolean } {
  if (state.allowClose) return { prevented: false, sentRequest: false };

  if (!rendererAlive) {
    state.allowClose = true;
    return { prevented: false, sentRequest: false };
  }

  if (state.closeRequestPending) {
    return { prevented: true, sentRequest: false };
  }

  state.closeRequestPending = true;
  return { prevented: true, sentRequest: true };
}

/** Обработчик window:confirmClose из рендерера. */
function onConfirmClose(state: CloseState, proceed: boolean): { closed: boolean } {
  state.closeRequestPending = false;
  if (!proceed) {
    state.allowClose = false;
    return { closed: false };
  }
  state.allowClose = true;
  return { closed: true };
}

/* ============================================================
   Проверки
   ============================================================ */

const failures: string[] = [];
let passed = 0;

function check(name: string, actual: unknown, expected: unknown) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) passed++;
  else failures.push(`${name}\n     ожидалось: ${e}\n     получено:   ${a}`);
}

/* ---- 1. Первое закрытие перехватывается и спрашивает рендерер ---- */
{
  const s = newState();
  const r = onWindowClose(s, true);
  check('закрытие перехвачено', r.prevented, true);
  check('запрос отправлен в рендерер', r.sentRequest, true);
  check('окно ещё не закрыто', s.allowClose, false);
}

/* ---- 2. Отмена: приложение продолжает работу ---- */
{
  const s = newState();
  onWindowClose(s, true);
  const r = onConfirmClose(s, false);
  check('отмена не закрывает окно', r.closed, false);
  check('флаг закрытия сброшен', s.allowClose, false);
  check('ожидание ответа сброшено', s.closeRequestPending, false);

  // Повторная попытка закрытия снова должна спросить
  const again = onWindowClose(s, true);
  check('после отмены закрытие снова перехвачено', again.prevented, true);
  check('после отмены запрос отправлен повторно', again.sentRequest, true);
}

/* ---- 3. Подтверждение: окно закрывается ---- */
{
  const s = newState();
  onWindowClose(s, true);
  const r = onConfirmClose(s, true);
  check('подтверждение закрывает окно', r.closed, true);
  check('разрешение на закрытие установлено', s.allowClose, true);

  // Следующее событие close проходит без перехвата
  const final = onWindowClose(s, true);
  check('повторное закрытие не перехватывается', final.prevented, false);
}

/* ---- 4. Повторные клики по «X» не открывают диалог дважды ---- */
{
  const s = newState();
  const first = onWindowClose(s, true);
  const second = onWindowClose(s, true);
  const third = onWindowClose(s, true);
  check('первый клик спрашивает', first.sentRequest, true);
  check('второй клик не дублирует диалог', second.sentRequest, false);
  check('третий клик не дублирует диалог', third.sentRequest, false);
  check('все клики отменяют закрытие', [first.prevented, second.prevented, third.prevented], [
    true,
    true,
    true,
  ]);
}

/* ---- 5. Рендерер мёртв — окно закрывается без диалога ---- */
{
  const s = newState();
  const r = onWindowClose(s, false);
  check('без рендерера закрытие не блокируется', r.prevented, false);
  check('без рендерера запрос не отправляется', r.sentRequest, false);
  check('без рендерера закрытие разрешено', s.allowClose, true);
}

/* ---- 6. Полный сценарий: правка → закрытие → сохранение → выход ---- */
{
  const s = newState();
  // Пользователь нажал X с несохранёнными файлами
  const clicked = onWindowClose(s, true);
  check('X с несохранёнными файлами перехвачен', clicked.prevented, true);
  // Диалог: «Сохранить всё и выйти» → подтверждение
  const confirmed = onConfirmClose(s, true);
  check('выход после сохранения разрешён', confirmed.closed, true);
  check('окно закрывается', onWindowClose(s, true).prevented, false);
}

/* ---- 7. Сценарий потери данных больше не воспроизводится ---- */
{
  // Раньше: closeWindow() закрывал окно мгновенно, изменения терялись.
  // Теперь любое закрытие проходит через перехват.
  const s = newState();
  const viaTitleBar = onWindowClose(s, true); // кнопка X в заголовке
  check('кнопка в заголовке тоже перехватывается', viaTitleBar.prevented, true);
  check('и запрашивает подтверждение', viaTitleBar.sentRequest, true);
  check('изменения не теряются молча', s.allowClose, false);
}

console.log('='.repeat(72));
console.log('  ПРОВЕРКА ПОДТВЕРЖДЕНИЯ ВЫХОДА');
console.log('='.repeat(72));
console.log(`  Пройдено: ${passed}, провалов: ${failures.length}`);
if (failures.length > 0) {
  console.log('-'.repeat(72));
  for (const f of failures) console.log(`   - ${f}`);
}
console.log('='.repeat(72));
console.log(failures.length === 0 ? '  ИТОГ: ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ' : '  ИТОГ: ЕСТЬ ПРОВАЛЫ');
process.exit(failures.length === 0 ? 0 : 1);
