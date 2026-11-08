/**
 * Unit-проверки documentModel в настоящем DOM приложения.
 *
 * Приложение открывается с ?modelTest=1 и само выставляет window.__documentModel.
 * Сценарии работают только с HTML-строкой, поэтому не зависят от редактора,
 * но используют тот же собранный модуль, что и прод.
 *
 * Запуск: electron.exe --no-sandbox dist-electron/model-check.js
 */
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { registerCommonCheckIpc } = require('./common-ipc');

function resolveProjectRoot() {
  let dir = __dirname;
  for (let i = 0; i < 4; i++) {
    if (fs.existsSync(path.join(dir, 'package.json')) && fs.existsSync(path.join(dir, 'dist'))) {
      return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  return path.join(__dirname, '..');
}

const ROOT = resolveProjectRoot();
const SANDBOX = process.env.WS_CHECK_SANDBOX || path.join(ROOT, 'test-workspace-model');
const OUT = path.join(ROOT, 'model-check.json');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function registerIpc() {
  registerCommonCheckIpc({ sandbox: SANDBOX });
  ipcMain.handle('window:minimize', () => undefined);
  ipcMain.handle('window:maximize', () => false);
  ipcMain.handle('window:isMaximized', () => false);
  ipcMain.handle('window:close', () => undefined);
  ipcMain.handle('fs:getCurrentWorkspace', () => SANDBOX);
  ipcMain.handle('fs:setWorkspaceRoot', () => ({ success: true }));
  ipcMain.handle('fs:openFolderDialog', () => SANDBOX);
  ipcMain.handle('fs:openFileDialog', () => null);
  ipcMain.handle('fs:saveFileDialog', () => null);
  ipcMain.handle('fs:showInExplorer', () => ({ success: true }));
  ipcMain.handle('fs:listDirectory', async () => ({ success: true, items: [] }));
  ipcMain.handle('fs:readFile', async () => ({ success: false, error: 'not used' }));
  ipcMain.handle('fs:writeFile', async () => ({ success: true }));
  ipcMain.handle('fs:createFile', async () => ({ success: true }));
  ipcMain.handle('fs:createDirectory', async () => ({ success: true }));
  ipcMain.handle('fs:deletePath', async () => ({ success: true }));
  ipcMain.handle('fs:renamePath', async () => ({ success: true }));
  ipcMain.handle('fs:searchWorkspace', async () => ({ success: true, results: [] }));
  ipcMain.handle('ai:getSettings', () => ({ configured: false }));
  ipcMain.handle('app:getVersion', () => '1.0.0');
}

async function main() {
  fs.mkdirSync(SANDBOX, { recursive: true });
  registerIpc();

  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    show: false,
    webPreferences: {
      preload: path.join(ROOT, 'dist-electron', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  const out: any = { steps: [], errors: [] };
  win.webContents.on('console-message', (_event, level, message) => {
    out.errors.push(`${level}: ${message}`);
  });

  try {
    await win.loadFile(path.join(ROOT, 'dist', 'index.html'), { query: { modelTest: '1' } });
    await sleep(2500);

    const ready = await win.webContents.executeJavaScript('!!window.__documentModel');
    if (!ready) throw new Error('window.__documentModel не появился: тестовый мост не собран');

    const steps = await win.webContents.executeJavaScript(`
      (async () => {
        const M = window.__documentModel;
        const run = (html, tool, args) => M.applyDocumentTool(html, tool, args || {});
        const steps = [];
        const add = (name, value) => steps.push({ name, ...value });

        /* 1. replace_document тем же текстом отклоняется */
        const same = run('<p>Текст документа</p>', 'replace_document', { html: '<p>Текст документа</p>' });
        add('noopReplace', { rejected: same.ok === false, message: /не изменилось/.test(same.message) });

        /* 2. Отчёт ассистента не попадает в документ */
        const noisy = run(
          '<p>Старый текст</p>',
          'replace_document',
          { html: '<p>Оформил документ в деловом стиле. Вот итоговая версия:</p><p>Новый текст документа</p><p>Документ Test.docx обновлён — сохранить сейчас?</p>' }
        );
        add('reportCleanup', {
          ok: noisy.ok === true,
          bodyKept: (noisy.html || '').includes('Новый текст документа'),
          reportGone: !(noisy.html || '').includes('Оформил документ'),
          offerGone: !(noisy.html || '').includes('сохранить сейчас'),
        });

        /* 3. replace_block меняет только указанный блок */
        const block = run('<p data-bid="b1">Первый</p><p data-bid="b2">Второй</p>', 'replace_block', {
          block_id: 'b2', html: '<p>Заменённый</p>',
        });
        add('replaceBlock', {
          ok: block.ok === true,
          firstKept: (block.html || '').includes('Первый'),
          secondReplaced: (block.html || '').includes('Заменённый') && !(block.html || '').includes('Второй'),
        });

        /* 4. insert_after вставляет после блока, не задевая его */
        const inserted = run('<p data-bid="b1">Первый</p><p data-bid="b2">Второй</p>', 'insert_after', {
          block_id: 'b1', html: '<p>Вставленный</p>',
        });
        const insertedHtml = inserted.html || '';
        add('insertAfter', {
          ok: inserted.ok === true,
          order: insertedHtml.indexOf('Первый') < insertedHtml.indexOf('Вставленный') &&
                 insertedHtml.indexOf('Вставленный') < insertedHtml.indexOf('Второй'),
        });

        /* 4b. insert_before и append_to_block */
        const before = run('<p data-bid="b1">Первый</p><p data-bid="b2">Второй</p>', 'insert_before', {
          block_id: 'b2', html: '<p>Перед вторым</p>',
        });
        const beforeHtml = before.html || '';
        add('insertBefore', {
          ok: before.ok === true,
          order:
            beforeHtml.indexOf('Первый') < beforeHtml.indexOf('Перед вторым') &&
            beforeHtml.indexOf('Перед вторым') < beforeHtml.indexOf('Второй'),
        });

        const appended = run('<p data-bid="b1">Начало</p>', 'append_to_block', {
          block_id: 'b1', html: '<strong> конец</strong>',
        });
        const appendedProbe = document.createElement('div');
        appendedProbe.innerHTML = appended.html || '';
        add('appendToBlock', {
          ok: appended.ok === true,
          text: appendedProbe.textContent === 'Начало конец',
          bold: (appended.html || '').includes('<strong>'),
        });

        const appendedList = run('<ul><li data-bid="b1">Первый</li></ul>', 'append_to_block', {
          block_id: 'b1', html: '<li>Второй</li>',
        });
        add('appendToList', {
          ok: appendedList.ok === true,
          items: (appendedList.html || '').includes('Второй'),
        });

        /* 4c. Перемещение, дублирование, разделение и слияние блоков */
        const moved = run(
          '<p data-bid="b1">Первый</p><p data-bid="b2">Второй</p><p data-bid="b3" style="text-align:center">Третий</p>',
          'move_block',
          { block_id: 'b3', before_id: 'b1' }
        );
        const movedHtml = moved.html || '';
        add('moveBlock', {
          ok: moved.ok === true,
          order:
            movedHtml.indexOf('Третий') < movedHtml.indexOf('Первый') &&
            movedHtml.indexOf('Первый') < movedHtml.indexOf('Второй'),
          styleKept: movedHtml.includes('text-align:center'),
        });

        const duplicated = run('<p data-bid="b1">Копия</p>', 'duplicate_block', { block_id: 'b1' });
        const duplicatedProbe = document.createElement('div');
        duplicatedProbe.innerHTML = duplicated.html || '';
        add('duplicateBlock', {
          ok: duplicated.ok === true,
          count: duplicatedProbe.children.length === 2,
          distinctIds: new Set(
            [...duplicatedProbe.querySelectorAll('[data-bid]')].map((node) =>
              node.getAttribute('data-bid')
            )
          ).size === 2,
        });

        const split = run('<p data-bid="b1">Первое предложение. Второе.</p>', 'split_block', {
          block_id: 'b1', offset: 19,
        });
        const splitProbe = document.createElement('div');
        splitProbe.innerHTML = split.html || '';
        add('splitBlock', {
          ok: split.ok === true,
          twoBlocks: splitProbe.children.length === 2,
          textPreserved: splitProbe.textContent === 'Первое предложение. Второе.',
        });

        const merged = run('<p data-bid="b1">Первое</p><p data-bid="b2">Второе</p>', 'merge_blocks', {
          block_id: 'b1',
        });
        const mergedProbe = document.createElement('div');
        mergedProbe.innerHTML = merged.html || '';
        add('mergeBlocks', {
          ok: merged.ok === true,
          oneBlock: mergedProbe.children.length === 1,
          text: mergedProbe.textContent === 'Первое Второе',
        });

        /* 4d. Таблицы: чтение и структурные правки */
        const tableHtml =
          '<table><tbody><tr><td> A </td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table>';
        const tableInfo = run(tableHtml, 'table_read', { block_id: 'b1' });
        add('tableRead', {
          ok: tableInfo.ok === true,
          rows: !!tableInfo.meta && tableInfo.meta.rows === 2,
          cols: !!tableInfo.meta && tableInfo.meta.cols === 2,
          cells: !!tableInfo.structure && tableInfo.structure.includes('c1="B"'),
        });

        const rowInserted = run(tableHtml, 'table_insert_row', {
          block_id: 'b1', index: 1, cells: ['X', 'Y'],
        });
        const rowProbe = document.createElement('div');
        rowProbe.innerHTML = rowInserted.html || '';
        add('tableInsertRow', {
          ok: rowInserted.ok === true,
          rows: rowProbe.querySelectorAll('tr').length === 3,
          cells: (rowProbe.querySelectorAll('tr')[1]?.textContent || '').includes('X'),
        });

        const rowDeleted = run(tableHtml, 'table_delete_row', { block_id: 'b1', index: 0 });
        const rowDelProbe = document.createElement('div');
        rowDelProbe.innerHTML = rowDeleted.html || '';
        add('tableDeleteRow', {
          ok: rowDeleted.ok === true,
          rows: rowDelProbe.querySelectorAll('tr').length === 1,
          kept: (rowDelProbe.textContent || '').includes('C'),
        });

        const colInserted = run(tableHtml, 'table_insert_column', { block_id: 'b1', index: 1 });
        const colProbe = document.createElement('div');
        colProbe.innerHTML = colInserted.html || '';
        add('tableInsertColumn', {
          ok: colInserted.ok === true,
          cols: [...colProbe.querySelectorAll('tr')].every((row) => row.children.length === 3),
        });

        const colDeleted = run(tableHtml, 'table_delete_column', { block_id: 'b1', index: 0 });
        const colDelProbe = document.createElement('div');
        colDelProbe.innerHTML = colDeleted.html || '';
        add('tableDeleteColumn', {
          ok: colDeleted.ok === true,
          cols: [...colDelProbe.querySelectorAll('tr')].every((row) => row.children.length === 1),
          kept: (colDelProbe.textContent || '').includes('B'),
        });

        const cellSet = run(tableHtml, 'table_set_cell', {
          block_id: 'b1', row: 0, col: 1, html: '<p>Новое</p>',
        });
        add('tableSetCell', {
          ok: cellSet.ok === true,
          changed: (cellSet.html || '').includes('Новое'),
          otherKept: (cellSet.html || '').includes('A') && (cellSet.html || '').includes('C'),
        });

        const mergedCells = run(tableHtml, 'table_merge_cells', {
          block_id: 'b1', row: 0, col: 0, colspan: 2,
        });
        const mergeProbe = document.createElement('div');
        mergeProbe.innerHTML = mergedCells.html || '';
        add('tableMergeCells', {
          ok: mergedCells.ok === true,
          colspan: mergeProbe.querySelector('td')?.getAttribute('colspan') === '2',
          firstRowCells: mergeProbe.querySelectorAll('tr')[0]?.children.length === 1,
        });

        const unmergedCells = run(
          '<table><tbody><tr><td colspan="2">Объединено</td></tr></tbody></table>',
          'table_unmerge_cells',
          { block_id: 'b1', row: 0, col: 0 }
        );
        const unmergeProbe = document.createElement('div');
        unmergeProbe.innerHTML = unmergedCells.html || '';
        add('tableUnmergeCells', {
          ok: unmergedCells.ok === true,
          cells: unmergeProbe.querySelectorAll('tr')[0]?.children.length === 2,
          colspanGone: !unmergeProbe.querySelector('td')?.getAttribute('colspan'),
        });

        const columnWidth = run(tableHtml, 'table_set_column_width', {
          block_id: 'b1',
          col: 0,
          width_pt: 100,
        });
        add('tableColumnWidth', {
          ok: columnWidth.ok === true,
          width: /width:\\s*100pt/.test(columnWidth.html || ''),
        });

        const tableBorder = run(tableHtml, 'table_set_borders', {
          block_id: 'b1',
          color: '#FF0000',
          size_pt: 2,
          style: 'dashed',
        });
        const borderProbe = document.createElement('div');
        borderProbe.innerHTML = tableBorder.html || '';
        add('tableBorders', {
          ok: tableBorder.ok === true,
          spec: borderProbe.querySelector('table')?.getAttribute('data-table-border') === '2|dashed|#FF0000',
          css: /border:\\s*2pt dashed/.test(tableBorder.html || ''),
        });

        const pageField = run('<p data-bid="b1">Текст</p>', 'insert_field', {
          block_id: 'b1',
          kind: 'PAGE',
          at: 'after',
        });
        const fieldProbe = document.createElement('div');
        fieldProbe.innerHTML = pageField.html || '';
        add('pageField', {
          ok: pageField.ok === true,
          field: fieldProbe.querySelector('[data-field="PAGE"]')?.textContent === '1',
        });

        const bookmark = run('<p data-bid="b1">Текст закладки</p>', 'insert_bookmark', {
          block_id: 'b1',
          name: 'Метка',
          find: 'закладки',
        });
        const bookmarkProbe = document.createElement('div');
        bookmarkProbe.innerHTML = bookmark.html || '';
        add('bookmark', {
          ok: bookmark.ok === true,
          marked: Boolean(bookmarkProbe.querySelector('[data-bookmark="Метка"]')),
          textKept: bookmarkProbe.textContent === 'Текст закладки',
        });

        const customStyle = run('<p data-bid="b1">Текст стиля</p>', 'define_style', {
          block_id: 'b1',
          name: 'Врезка',
          style_id: 'CustomCheck',
          font: 'Georgia',
          size_pt: 14,
          bold: true,
          color: '#112233',
          align: 'center',
        });
        const customProbe = document.createElement('div');
        customProbe.innerHTML = customStyle.html || '';
        const customBlock = customProbe.querySelector('[data-style="CustomCheck"]');
        add('customStyle', {
          ok: customStyle.ok === true,
          applied: Boolean(customBlock),
          font: /font-family:\\s*"?Georgia"?/.test(customStyle.html || ''),
          size: /font-size:\\s*14pt/.test(customStyle.html || ''),
          align: /text-align:\\s*center/.test(customStyle.html || ''),
        });

        const sectionRemoved = run(
          '<p data-bid="b1">Первый</p><p data-section-break="1"><br></p><p data-bid="b3">Второй</p>',
          'delete_section_break',
          {}
        );
        add('deleteSectionBreak', {
          ok: sectionRemoved.ok === true,
          gone: !(sectionRemoved.html || '').includes('data-section-break'),
          textKept:
            (sectionRemoved.html || '').includes('Первый') &&
            (sectionRemoved.html || '').includes('Второй'),
        });

        const commented = run('<p data-bid="b1">Текст для комментария</p>', 'add_comment', {
          block_id: 'b1',
          find: 'комментария',
          text: 'Проверить формулировку',
          author: 'AI',
        });
        const commentProbe = document.createElement('div');
        commentProbe.innerHTML = commented.html || '';
        add('addComment', {
          ok: commented.ok === true,
          marked: Boolean(commentProbe.querySelector('[data-comment-id]')),
          text: commentProbe.querySelector('[data-comment-text]')
            ?.getAttribute('data-comment-text') === 'Проверить формулировку',
        });

        const footnoted = run('<p data-bid="b1">Текст со сноской</p>', 'add_footnote', {
          block_id: 'b1',
          find: 'сноской',
          text: 'Источник сноски',
        });
        const footnoteProbe = document.createElement('div');
        footnoteProbe.innerHTML = footnoted.html || '';
        add('addFootnote', {
          ok: footnoted.ok === true,
          marked: Boolean(footnoteProbe.querySelector('[data-footnote-id]')),
          text:
            footnoteProbe
              .querySelector('[data-footnote-id]')
              ?.getAttribute('data-footnote-text') === 'Источник сноски',
          marker: Boolean(footnoteProbe.querySelector('sup[data-footnote-ref]')),
        });

        const refreshed = run(
          '<p data-field="TOC">Оглавление</p><h1>Глава</h1><h2>Раздел</h2>',
          'refresh_toc',
          {}
        );
        add('refreshToc', {
          ok: refreshed.ok === true,
          heading: (refreshed.html || '').includes('Глава'),
          subheading: (refreshed.html || '').includes('Раздел'),
          updated: (refreshed.html || '').includes('data-toc-updated'),
        });

        const changeReport = run(
          '<p>Старый текст резюме</p>',
          'replace_document',
          {
            html:
              '<p>Меняю должность на «Бухгалтер» и перерабатываю всё содержание резюме под эту профессию.</p>' +
              '<p>Все разделы обновлены под должность бухгалтера:</p>' +
              '<ul><li>Заголовок и шапка → «Резюме: Бухгалтер», должность в строке контактов.</li>' +
              '<li>Цель → замещение должности бухгалтера.</li>' +
              '<li>Опыт → обе позиции переименованы в «бухгалтер».</li>' +
              '<li>Образование → финансово-экономический вуз.</li></ul>' +
              '<p>Обратите внимание: остались технические детали прошлой версии.</p>' +
              '<p>Файл изменён, но не сохранён — сохранить?</p>',
          }
        );
        add('changeReportRejected', {
          rejected: changeReport.ok === false,
          message: /отч[её]т|очистк|служебн/i.test(changeReport.message || ''),
        });

        /* 4e. Изображения: вставка, размер/alt, удаление */
        const png =
          'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lK3hWQAAAABJRU5ErkJggg==';
        const imageInserted = run('<p data-bid="b1">До</p>', 'insert_image', {
          src: png, block_id: 'b1', width: 100, height: 50, alt: 'Тест',
        });
        const imageHtml = imageInserted.html || '';
        add('insertImage', {
          ok: imageInserted.ok === true,
          hasImage: imageHtml.includes('<img') && imageHtml.includes('data:image/png'),
          sized: imageHtml.includes('width:100px') && imageHtml.includes('height:50px'),
          alt: imageHtml.includes('alt="Тест"'),
        });

        const imageUpdated = run(imageHtml, 'update_image', {
          block_id: 'b2', width: 200, alt: 'Новый alt', align: 'center',
        });
        add('updateImage', {
          ok: imageUpdated.ok === true,
          resized: /width:\\s*200px/.test(imageUpdated.html || ''),
          alt: (imageUpdated.html || '').includes('alt="Новый alt"'),
          centered: (imageUpdated.html || '').includes('text-align: center'),
        });

        const imageDeleted = run(imageHtml, 'delete_image', { block_id: 'b2' });
        add('deleteImage', {
          ok: imageDeleted.ok === true,
          removed: !(imageDeleted.html || '').includes('<img'),
        });

        /* 4f. Разрыв страницы */
        const pageBreak = run('<p data-bid="b1">Текст</p>', 'insert_page_break', {
          block_id: 'b1',
        });
        add('insertPageBreak', {
          ok: pageBreak.ok === true,
          marked: (pageBreak.html || '').includes('data-page-break'),
          afterText:
            (pageBreak.html || '').indexOf('Текст') <
            (pageBreak.html || '').indexOf('data-page-break'),
        });

        /* 4f2. Поле оглавления */
        const toc = run('<p data-bid="b1">Первый</p><p data-bid="b2">Второй</p>', 'insert_toc', {
          block_id: 'b1',
        });
        const tocHtml = toc.html || '';
        add('insertToc', {
          ok: toc.ok === true,
          field: tocHtml.includes('data-field="TOC"'),
          order: tocHtml.indexOf('Первый') < tocHtml.indexOf('data-field="TOC"'),
        });

        const sectionBreak = run('<p data-bid="b1">Первый</p><p data-bid="b2">Второй</p>', 'insert_section_break', {
          block_id: 'b1',
        });
        add('insertSectionBreak', {
          ok: sectionBreak.ok === true,
          marked: (sectionBreak.html || '').includes('data-section-break'),
          beforeSecond:
            (sectionBreak.html || '').indexOf('data-section-break') <
            (sectionBreak.html || '').indexOf('Второй'),
        });

        /* 4g. Формат абзаца: интервалы и запреты разрыва */
        const paragraphFormat = run('<p data-bid="b1">Абзац</p>', 'set_paragraph_format', {
          block_id: 'b1',
          line_spacing: 1.5,
          space_before_pt: 6,
          space_after_pt: 4,
          keep_with_next: true,
          keep_together: true,
          page_break_before: true,
        });
        const paragraphHtml = paragraphFormat.html || '';
        add('paragraphFormat', {
          ok: paragraphFormat.ok === true,
          line: /line-height:\s*1\.5/.test(paragraphHtml),
          before: /margin-top:\s*6pt/.test(paragraphHtml),
          after: /margin-bottom:\s*4pt/.test(paragraphHtml),
          keep: /break-after:\s*avoid/.test(paragraphHtml) && /break-inside:\s*avoid/.test(paragraphHtml),
          pageBreakBefore: /break-before:\s*page/.test(paragraphHtml),
        });

        const widow = run('<p data-bid="b1">Текст</p>', 'set_paragraph_format', {
          block_id: 'b1',
          widow_control: false,
        });
        add('widowControl', {
          ok: widow.ok === true,
          disabled: (widow.html || '').includes('data-widow="off"'),
        });

        /* 4h. Заливка и рамка блока */
        const styled = run('<p data-bid="b1">Блок</p>', 'format', {
          block_id: 'b1',
          fill: '#FFFF00',
          border: true,
          border_color: '#000000',
          border_width_pt: 2,
          border_style: 'dashed',
        });
        const styledHtml = styled.html || '';
        add('blockFillAndBorder', {
          ok: styled.ok === true,
          fill: /background-color:\s*#FFFF00/i.test(styledHtml),
          border: /border:\s*2pt dashed #000000/i.test(styledHtml),
          textKept: styledHtml.includes('Блок'),
        });

        const tabs = run('<p data-bid="b1">Текст</p>', 'set_tabs', {
          block_id: 'b1',
          positions: '24,48',
        });
        add('setTabs', {
          ok: tabs.ok === true,
          stops: (tabs.html || '').includes('data-tabs="24,48"'),
        });

        const paragraphStyle = run('<p data-bid="b1">Текст</p>', 'set_paragraph_style', {
          block_id: 'b1',
          style: 'Title',
        });
        add('paragraphStyle', {
          ok: paragraphStyle.ok === true,
          style: (paragraphStyle.html || '').includes('data-style="Title"'),
        });

        /* 4i. Гиперссылка на фрагменте */
        const linked = run('<p data-bid="b1">Наш сайт здесь</p>', 'format', {
          block_id: 'b1',
          find: 'сайт',
          link: 'https://example.com',
        });
        const linkedHtml = linked.html || '';
        add('formatLink', {
          ok: linked.ok === true,
          href: linkedHtml.includes('href="https://example.com"'),
          textKept: linkedHtml.includes('Наш') && linkedHtml.includes('здесь'),
        });

        /* 5. delete_block удаляет только указанный блок */
        const deleted = run('<p data-bid="b1">Первый</p><p data-bid="b2">Второй</p>', 'delete_block', { block_id: 'b1' });
        add('deleteBlock', {
          ok: deleted.ok === true,
          removed: !(deleted.html || '').includes('Первый'),
          kept: (deleted.html || '').includes('Второй'),
        });

        /* 6. find_replace через несколько runs */
        const crossRun = run(
          '<p data-bid="b1"><strong>Иван</strong> <em>Иванов</em></p>',
          'find_replace',
          { find: 'Иван Иванов', replace: 'Пётр Сидоров' }
        );
        add('crossRunReplace', {
          ok: crossRun.ok === true,
          replaced: (crossRun.html || '').includes('Пётр Сидоров'),
          oldGone: !(crossRun.html || '').includes('Иван'),
        });

        /* 7. find_replace с NBSP */
        const nbsp = run('<p data-bid="b1">Иван\u00a0Иванов</p>', 'find_replace', {
          find: 'Иван Иванов', replace: 'Пётр Сидоров',
        });
        add('nbspReplace', { ok: nbsp.ok === true, replaced: (nbsp.html || '').includes('Пётр Сидоров') });

        /* 8. format форматирует только найденный фрагмент */
        const formatted = run('<p data-bid="b1">Половина текста здесь</p>', 'format', {
          find: 'текста', bold: true,
        });
        const formattedHtml = formatted.html || '';
        add('formatRange', {
          ok: formatted.ok === true,
          boldWrapped: /<strong>текста<\\/strong>/.test(formattedHtml),
          prefixIntact: formattedHtml.includes('Половина'),
          suffixIntact: formattedHtml.includes('здесь'),
        });

        /* 8b. Форматирование половины слова не задевает соседние буквы */
        const halfWord = run('<p data-bid="b1">Половина текста здесь</p>', 'format', {
          find: 'кста', bold: true,
        });
        const halfHtml = halfWord.html || '';
        const halfProbe = document.createElement('div');
        halfProbe.innerHTML = halfHtml;
        add('formatHalfWord', {
          ok: halfWord.ok === true,
          marked: /те<strong>кста<\\/strong>/.test(halfHtml),
          textPreserved: halfProbe.textContent === 'Половина текста здесь',
        });

        const crossFormat = run(
          '<p data-bid="b1"><strong>Иван</strong> Иванов</p>',
          'format',
          { find: 'Иван Иванов', italic: true }
        );
        const crossProbe = document.createElement('div');
        crossProbe.innerHTML = crossFormat.html || '';
        add('formatCrossRun', {
          ok: crossFormat.ok === true,
          italic: (crossFormat.html || '').includes('<em'),
          textPreserved: crossProbe.textContent === 'Иван Иванов',
        });

        const colored = run('<p data-bid="b1">Цветной текст</p>', 'format', {
          find: 'текст', color: '#C00000', highlight: 'yellow',
        });
        const coloredHtml = colored.html || '';
        add('formatColorAndHighlight', {
          ok: colored.ok === true,
          color: coloredHtml.includes('color:#C00000'),
          highlight: coloredHtml.includes('background-color:yellow'),
        });

        /* 9. page_setup возвращает параметры страницы */
        const page = run('<p>Текст</p>', 'page_setup', {
          margins: { top: 3, right: 1.5, bottom: 2, left: 2.5 },
          fontFamily: 'Times New Roman',
          size_pt: 14,
        });
        add('pageSetup', {
          ok: page.ok === true,
          pageReturned: !!page.page && page.page.fontFamily === 'Times New Roman' && page.page.fontSizePt === 14,
        });

        /* 10. Read-only объект сохраняется при полной переработке */
        const readonly = run(
          '<p data-docx-readonly="1" data-docx-kind="sdt">[Элемент Word — только просмотр]</p><p data-bid="b1">Старый текст</p>',
          'replace_document',
          { html: '<p>Новый документ</p>' }
        );
        add('readonlyPreserved', {
          ok: readonly.ok === true,
          newBody: (readonly.html || '').includes('Новый документ'),
          readonlyKept: (readonly.html || '').includes('data-docx-readonly'),
        });

        /* 11. Read-only блок нельзя заменить или удалить */
        const readonlyReplace = run(
          '<p data-docx-readonly="1">[Объект — только просмотр]</p><p data-bid="b2">Текст</p>',
          'replace_block',
          { block_id: 'b1', html: '<p>Подмена</p>' }
        );
        const readonlyDelete = run(
          '<p data-docx-readonly="1">[Объект — только просмотр]</p><p data-bid="b2">Текст</p>',
          'delete_block',
          { block_id: 'b1' }
        );
        add('readonlyGuarded', {
          replaceRejected: readonlyReplace.ok === false,
          deleteRejected: readonlyDelete.ok === false,
        });

        /* 12. Вставка только отчёта отклоняется */
        const reportOnly = run('<p data-bid="b1">Текст</p>', 'insert_after', {
          block_id: 'b1',
          html: '<p>Оформил документ. Вот итоговая версия:</p>',
        });
        add('reportOnlyRejected', { rejected: reportOnly.ok === false });

        /* 13. Пагинация outline */
        const outline = run(
          '<p data-bid="b1">Первый блок</p><p data-bid="b2">Второй блок</p><p data-bid="b3">Третий блок</p>',
          'get_outline',
          { offset: 1, limit: 1 }
        );
        add('outlinePage', {
          ok: outline.ok === true,
          secondOnly: (outline.structure || '').includes('Второй блок') &&
                      !(outline.structure || '').includes('Первый блок') &&
                      !(outline.structure || '').includes('Третий блок'),
          next: !!outline.meta && outline.meta.nextOffset === 2,
          total: !!outline.meta && outline.meta.total === 3,
        });

        /* 14. Полный HTML блока */
        const blockInfo = run('<p data-bid="b1">Первый</p><h2 data-bid="b2">Второй</h2>', 'get_block', {
          block_id: 'b2',
        });
        add('getBlock', {
          ok: blockInfo.ok === true,
          content: (blockInfo.content || '').includes('<h2') && (blockInfo.content || '').includes('Второй'),
        });

        /* 15. Диапазон внутри блока */
        const rangeInfo = run('<p data-bid="b1">Привет мир</p>', 'get_range', {
          block_id: 'b1', start: 0, end: 6,
        });
        add('getRange', {
          ok: rangeInfo.ok === true,
          content: rangeInfo.content === 'Привет',
          meta: !!rangeInfo.meta && rangeInfo.meta.start === 0 && rangeInfo.meta.end === 6,
        });

        /* 16. Поиск с точными смещениями через runs */
        const search = run(
          '<p data-bid="b1"><strong>Иван</strong> Иванов</p><p data-bid="b2">Другой Иван</p>',
          'search_document',
          { query: 'Иванов' }
        );
        add('searchDocument', {
          ok: search.ok === true,
          matches: !!search.matches && search.matches.length === 1,
          block: !!search.matches && search.matches[0].block_id === 'b1',
          offset: !!search.matches && search.matches[0].start === 5,
        });

        /* 17. Замена диапазона внутри абзаца */
        const rangeReplace = run('<p data-bid="b1">Привет мир</p>', 'replace_range', {
          block_id: 'b1', start: 7, end: 10, html: '<strong>друг</strong>',
        });
        add('replaceRange', {
          ok: rangeReplace.ok === true,
          replaced: (rangeReplace.html || '').includes('Привет <strong>друг</strong>'),
          oldGone: !(rangeReplace.html || '').includes('мир'),
        });

        /* 18. Удаление диапазона не задевает соседний текст */
        const rangeDelete = run('<p data-bid="b1">Привет мир</p>', 'delete_range', {
          block_id: 'b1', start: 6, end: 10,
        });
        add('deleteRange', {
          ok: rangeDelete.ok === true,
          text: (rangeDelete.html || '').includes('Привет'),
          deleted: !(rangeDelete.html || '').includes('мир'),
        });

        /* 19. Диапазон в read-only блоке запрещён */
        const rangeReadonly = run(
          '<p data-docx-readonly="1">[Объект — только просмотр]</p><p data-bid="b2">Текст</p>',
          'replace_range',
          { block_id: 'b1', start: 0, end: 1, html: 'X' }
        );
        add('rangeReadonlyGuard', { rejected: rangeReadonly.ok === false });

        /* 20. Путь другого файла не выходит за workspace */
        const P = window.__workspacePath;
        const root = 'C:/Builders/Workspace';
        const inside = P.resolveWorkspacePath(root, 'docs/Отчёт.docx');
        const absoluteInside = P.resolveWorkspacePath(root, 'C:/Builders/Workspace/notes.md');
        const outside = P.resolveWorkspacePath(root, '../Windows/win.ini');
        const absoluteOutside = P.resolveWorkspacePath(root, 'C:/Windows/win.ini');
        const empty = P.resolveWorkspacePath(root, '   ');
        add('workspacePathGuard', {
          inside: Boolean(inside && inside.replace(/\\\\/g, '/').toLowerCase() === 'c:/builders/workspace/docs/отчёт.docx'),
          absoluteInside: Boolean(absoluteInside && absoluteInside.replace(/\\\\/g, '/').toLowerCase().endsWith('notes.md')),
          outsideRejected: outside === null && absoluteOutside === null,
          emptyRejected: empty === null,
        });

        /* 20b. Структурные правки Markdown */
        const MD = window.__markdownOps;
        const mdHeading = MD.setMarkdownHeading('# Старый заголовок\\n\\nТекст', 'Текст', 2);
        add('markdownHeading', {
          ok: mdHeading.ok,
          heading: /^## Текст$/m.test(mdHeading.text),
          oldKept: mdHeading.text.includes('# Старый заголовок'),
        });
        const mdListed = MD.toggleMarkdownList('Пункт списка', 'Пункт', 'numbered');
        const mdUnlisted = MD.toggleMarkdownList(mdListed.text, 'Пункт', 'numbered');
        add('markdownList', {
          on: /^1\\. Пункт списка$/m.test(mdListed.text),
          off: /^Пункт списка$/m.test(mdUnlisted.text),
        });
        const mdLinked = MD.insertMarkdownLink('Сайт и всё', 'Сайт', 'Сайт', 'https://example.com');
        add('markdownLink', {
          ok: mdLinked.ok,
          linked: mdLinked.text.includes('[Сайт](https://example.com)'),
          rejected: MD.insertMarkdownLink('x', 'x', 'x', 'javascript:alert(1)').ok === false,
        });
        const mdTable = MD.insertMarkdownTable('Вступление', ['Имя', 'Роль'], [['Иван', 'Dev']]);
        add('markdownTable', {
          ok: mdTable.ok,
          header: mdTable.text.includes('| Имя | Роль |'),
          row: mdTable.text.includes('| Иван | Dev |'),
        });

        /* 21. Diff и история AI-транзакций */
        const H = window.__documentHistory;
        const diff = H.diffDocumentHtml(
          '<p data-docx-id="d1">Первый</p><p data-docx-id="d2">Второй</p>',
          '<p data-docx-id="d1">Первый изменён</p><p data-docx-id="d3">Третий</p>'
        );
        add('documentDiff', {
          changed: diff.changed === 1,
          added: diff.added === 1,
          removed: diff.removed === 1,
          blocks: diff.blocks.length === 3,
        });

        const history = new H.DocumentHistory(10);
        history.commit({ id: 't1', label: 'Правка 1', timestamp: 1, before: '<p>A</p>', after: '<p>B</p>' });
        history.commit({ id: 't2', label: 'Правка 2', timestamp: 2, before: '<p>B</p>', after: '<p>C</p>' });
        const undone = history.undo();
        const redoAvailable = history.canRedo();
        const redone = history.redo();
        history.undo();
        history.commit({ id: 't3', label: 'Правка 3', timestamp: 3, before: '<p>B</p>', after: '<p>D</p>' });
        add('transactionHistory', {
          undo: undone && undone.id === 't2',
          redo: redone && redone.id === 't2',
          redoWasAvailable: redoAvailable === true,
          futureTruncated: history.canRedo() === false,
          last: history.current && history.current.id === 't3',
        });

        /* 22. Реестр сессий документов */
        const S = window.__documentSession;
        S.resetDocumentSessions();
        const session = S.registerDocumentSession({
          id: 'tab-1',
          kind: S.kindForTab('word'),
          path: 'C:/work/report.docx',
          name: 'report.docx',
          writeAccess: true,
          tools: ['get_document', 'format'],
        });
        const updated = S.updateDocumentSession('tab-1', {
          capabilities: { write: false },
        });
        const listed = S.listDocumentSessions();
        S.removeDocumentSession('tab-1');
        add('documentSession', {
          kind: session.kind === 'word' && S.kindForTab('excel') === 'excel' &&
            S.kindForTab('powerpoint') === 'powerpoint' && S.kindForTab('markdown') === 'markdown',
          created: session.name === 'report.docx' && session.capabilities.read === true,
          updated: Boolean(updated) && updated.capabilities.write === false,
          listed: listed.length === 1 && listed[0].capabilities.tools.length === 2,
          removed: S.getDocumentSession('tab-1') === undefined,
        });

        const PPT = window.__pptx;
        const PPT_TPL = window.__pptxTemplate;
        const blank = PPT && PPT_TPL ? await PPT.parsePptx(PPT_TPL.BLANK_PPTX_BASE64) : null;
        add('blankPptx', {
          parsed: Boolean(blank),
          slides: Boolean(blank && blank.slides.length === 1),
          title: Boolean(blank && blank.slides[0]?.shapes.some((shape) => shape.text.includes('Новый слайд'))),
          layouts: Boolean(blank && blank.layouts.length > 0),
        });

        return steps;
      })()
    `);

    for (const step of steps) out.steps.push(step);
    const failed = out.steps.filter((step) =>
      Object.entries(step).some(([key, value]) => key !== 'name' && value !== true)
    );
    out.passed = failed.length === 0;
    out.failed = failed.map((step) => step.name);
  } catch (err) {
    out.fatal = String(err && err.stack ? err.stack : err);
  }

  fs.writeFileSync(OUT, JSON.stringify(out, null, 2), 'utf-8');
  win.destroy();
  app.quit();
}

app.whenReady().then(main).catch((err) => {
  fs.writeFileSync(OUT, JSON.stringify({ fatal: String(err) }, null, 2), 'utf-8');
  app.quit();
});
