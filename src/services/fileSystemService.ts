import { FileItem, FileType, SearchResult } from '../types';

/** Определяет категорию файла по расширению. */
export function getFileType(ext: string): FileType {
  const extension = ext.toLowerCase().replace(/^\./, '');
  switch (extension) {
    case 'xlsx':
    case 'xls':
    case 'xlsm':
    case 'csv':
    case 'tsv':
      return 'excel';
    case 'docx':
    case 'doc':
      return 'word';
    case 'pptx':
    case 'ppt':
      return 'powerpoint';
    case 'md':
    case 'markdown':
    case 'mdx':
      return 'markdown';
    case 'note':
    case 'rtf':
      return 'richtext';
    case 'png':
    case 'apng':
    case 'jpg':
    case 'jpeg':
    case 'jpe':
    case 'jfif':
    case 'pjpeg':
    case 'gif':
    case 'svg':
    case 'svgz':
    case 'webp':
    case 'avif':
    case 'ico':
    case 'cur':
    case 'bmp':
    case 'dib':
    case 'tif':
    case 'tiff':
    case 'heic':
    case 'heif':
    case 'jp2':
    case 'j2k':
    case 'jpf':
    case 'jpx':
    case 'wbmp':
    case 'xbm':
    case 'xpm':
    case 'dds':
    case 'exr':
    case 'hdr':
    case 'tga':
    case 'pcx':
    case 'ppm':
    case 'pgm':
    case 'pbm':
    case 'pnm':
    case 'psd':
    case 'emf':
    case 'wmf':
      return 'image';
    case 'pdf':
      return 'pdf';
    case 'mp4':
    case 'm4v':
    case 'mov':
    case 'webm':
    case 'mkv':
    case 'avi':
    case 'wmv':
    case 'flv':
    case 'f4v':
    case 'mpeg':
    case 'mpg':
    case 'mpe':
    case 'mp2':
    case 'm2v':
    case 'ts':
    case 'mts':
    case 'm2ts':
    case 'vob':
    case 'ogv':
    case 'ogg':
    case '3gp':
    case '3g2':
    case 'asf':
    case 'rm':
    case 'rmvb':
    case 'divx':
    case 'qt':
      return 'video';
    case 'js':
    case 'jsx':
    case 'mjs':
    case 'cjs':
    case 'ts':
    case 'tsx':
    case 'json':
    case 'html':
    case 'htm':
    case 'css':
    case 'scss':
    case 'less':
    case 'py':
    case 'sql':
    case 'sh':
    case 'bash':
    case 'bat':
    case 'cmd':
    case 'ps1':
    case 'yaml':
    case 'yml':
    case 'xml':
    case 'toml':
    case 'c':
    case 'cpp':
    case 'h':
    case 'hpp':
    case 'cs':
    case 'java':
    case 'go':
    case 'rs':
    case 'php':
    case 'rb':
    case 'lua':
      return 'code';
    case 'txt':
    case 'log':
    case 'ini':
    case 'cfg':
    case 'conf':
    case 'env':
    case 'properties':
    case 'gitignore':
    case 'dockerignore':
      return 'text';
    default:
      // Неизвестные расширения без точки (LICENSE, Makefile) — текстовые
      return extension ? 'binary' : 'text';
  }
}

/** Сопоставление расширения с идентификатором языка Monaco. */
export function getMonacoLanguage(ext: string): string {
  const map: Record<string, string> = {
    js: 'javascript',
    jsx: 'javascript',
    mjs: 'javascript',
    cjs: 'javascript',
    ts: 'typescript',
    tsx: 'typescript',
    json: 'json',
    html: 'html',
    htm: 'html',
    css: 'css',
    scss: 'scss',
    less: 'less',
    py: 'python',
    sql: 'sql',
    sh: 'shell',
    bash: 'shell',
    bat: 'bat',
    cmd: 'bat',
    ps1: 'powershell',
    yaml: 'yaml',
    yml: 'yaml',
    xml: 'xml',
    toml: 'ini',
    md: 'markdown',
    markdown: 'markdown',
    ini: 'ini',
    cfg: 'ini',
    conf: 'ini',
    env: 'ini',
    properties: 'ini',
    log: 'plaintext',
    txt: 'plaintext',
    c: 'c',
    cpp: 'cpp',
    h: 'cpp',
    hpp: 'cpp',
    cs: 'csharp',
    java: 'java',
    go: 'go',
    rs: 'rust',
    php: 'php',
    rb: 'ruby',
    lua: 'lua',
    dockerfile: 'dockerfile',
  };
  return map[ext.toLowerCase().replace(/^\./, '')] || 'plaintext';
}

/** Список поддерживаемых кодировок для переключателя в редакторе. */
export const SUPPORTED_ENCODINGS = [
  { id: 'utf-8', label: 'UTF-8' },
  { id: 'utf-16le', label: 'UTF-16 LE' },
  { id: 'windows-1251', label: 'Windows-1251 (ANSI)' },
  { id: 'windows-1252', label: 'Windows-1252' },
  { id: 'koi8-r', label: 'KOI8-R' },
  { id: 'iso-8859-5', label: 'ISO-8859-5' },
] as const;

export const fileSystemService = {
  isElectron(): boolean {
    return typeof window !== 'undefined' && Boolean(window.electronAPI);
  },

  async getCurrentWorkspace(): Promise<string> {
    if (this.isElectron()) return window.electronAPI!.getCurrentWorkspace();
    return '';
  },

  async setWorkspaceRoot(dirPath: string): Promise<void> {
    if (this.isElectron()) await window.electronAPI!.setWorkspaceRoot(dirPath);
  },

  async getFavorites(): Promise<{ path: string; name: string }[]> {
    if (!this.isElectron() || !window.electronAPI?.getFavorites) return [];
    return window.electronAPI.getFavorites();
  },

  async toggleFavorite(dirPath: string): Promise<{ path: string; name: string }[]> {
    if (!this.isElectron() || !window.electronAPI?.toggleFavorite) return [];
    const res = await window.electronAPI.toggleFavorite(dirPath);
    return res.favorites || [];
  },

  async openFolderDialog(): Promise<string | null> {
    if (this.isElectron()) return window.electronAPI!.openFolderDialog();
    return null;
  },

  async openFileDialog(filters?: { name: string; extensions: string[] }[]): Promise<string | null> {
    if (this.isElectron()) return window.electronAPI!.openFileDialog(filters);
    return null;
  },

  async saveFileDialog(
    defaultName: string,
    filters?: { name: string; extensions: string[] }[]
  ): Promise<string | null> {
    if (this.isElectron()) return window.electronAPI!.saveFileDialog(defaultName, filters);
    return null;
  },

  async showInExplorer(filePath: string): Promise<boolean> {
    if (!this.isElectron()) return false;
    const res = await window.electronAPI!.showInExplorer(filePath);
    return res.success;
  },

  async listDirectory(dirPath: string): Promise<FileItem[]> {
    if (!this.isElectron()) {
      throw new Error('Файловая система доступна только в настольном приложении WorkSpace');
    }
    const res = await window.electronAPI!.listDirectory(dirPath);
    if (!res.success) throw new Error(res.error || 'Ошибка чтения папки');
    return res.items.map((it) => ({
      ...it,
      type: it.isDirectory ? undefined : getFileType(it.extension),
    }));
  },

  async readFile(filePath: string, encoding = 'utf-8'): Promise<{ content: string; isBinary: boolean }> {
    if (!this.isElectron()) throw new Error('Чтение файлов доступно только в настольном приложении');
    const res = await window.electronAPI!.readFile(filePath, encoding);
    if (!res.success || res.content === undefined) {
      throw new Error(res.error || 'Ошибка чтения файла');
    }
    return { content: res.content, isBinary: Boolean(res.isBinary) };
  },

  async writeFile(filePath: string, content: string, encoding = 'utf-8'): Promise<void> {
    if (!this.isElectron()) throw new Error('Запись файлов доступна только в настольном приложении');
    const res = await window.electronAPI!.writeFile(filePath, content, encoding);
    if (!res.success) throw new Error(res.error || 'Ошибка записи файла');
  },

  async createFile(filePath: string, initialContent = '', encoding?: string): Promise<void> {
    if (!this.isElectron()) throw new Error('Создание файлов доступно только в настольном приложении');
    const res = await window.electronAPI!.createFile(filePath, initialContent, encoding);
    if (!res.success) throw new Error(res.error || 'Ошибка создания файла');
  },

  async createDirectory(dirPath: string): Promise<void> {
    if (!this.isElectron()) throw new Error('Создание папок доступно только в настольном приложении');
    const res = await window.electronAPI!.createDirectory(dirPath);
    if (!res.success) throw new Error(res.error || 'Ошибка создания папки');
  },

  async deletePath(targetPath: string): Promise<void> {
    if (!this.isElectron()) throw new Error('Удаление доступно только в настольном приложении');
    const res = await window.electronAPI!.deletePath(targetPath);
    if (!res.success) throw new Error(res.error || 'Ошибка удаления');
  },

  async renamePath(oldPath: string, newPath: string): Promise<void> {
    if (!this.isElectron()) throw new Error('Переименование доступно только в настольном приложении');
    const res = await window.electronAPI!.renamePath(oldPath, newPath);
    if (!res.success) throw new Error(res.error || 'Ошибка переименования');
  },

  async searchWorkspace(
    rootDir: string,
    query: string,
    searchContent = false,
    options?: { caseSensitive?: boolean; useRegex?: boolean; includeGlob?: string }
  ): Promise<SearchResult[]> {
    if (!this.isElectron()) return [];
    const res = await window.electronAPI!.searchWorkspace(rootDir, query, searchContent, options);
    if (!res.success) throw new Error(res.error || 'Ошибка поиска');
    return res.results;
  },
};

/* ============================================================
   Изображения: расширения и MIME-типы
   ============================================================ */

/** Все расширения, которые приложение считает изображениями. */
export const IMAGE_EXTENSIONS = [
  'png', 'apng', 'jpg', 'jpeg', 'jpe', 'jfif', 'pjpeg', 'gif', 'svg', 'svgz',
  'webp', 'avif', 'ico', 'cur', 'bmp', 'dib', 'tif', 'tiff', 'heic', 'heif',
  'jp2', 'j2k', 'jpf', 'jpx', 'wbmp', 'xbm', 'xpm', 'dds', 'exr', 'hdr',
  'tga', 'pcx', 'ppm', 'pgm', 'pbm', 'pnm', 'psd', 'emf', 'wmf',
] as const;

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  apng: 'image/apng',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  jfif: 'image/jpeg',
  pjpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  svgz: 'image/svg+xml',
  webp: 'image/webp',
  avif: 'image/avif',
  ico: 'image/x-icon',
  cur: 'image/x-icon',
  bmp: 'image/bmp',
  dib: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  heic: 'image/heic',
  heif: 'image/heif',
  jp2: 'image/jp2',
  j2k: 'image/jp2',
  jpf: 'image/jp2',
  jpx: 'image/jp2',
  wbmp: 'image/vnd.wap.wbmp',
  xbm: 'image/x-xbitmap',
  xpm: 'image/x-xpixmap',
  dds: 'image/vnd-ms.dds',
  exr: 'image/x-exr',
  hdr: 'image/vnd.radiance',
  tga: 'image/x-tga',
  pcx: 'image/x-pcx',
  ppm: 'image/x-portable-pixmap',
  pgm: 'image/x-portable-graymap',
  pbm: 'image/x-portable-bitmap',
  pnm: 'image/x-portable-anymap',
  psd: 'image/vnd.adobe.photoshop',
  emf: 'image/emf',
  wmf: 'image/wmf',
};

/** MIME-тип изображения по расширению (с разумным запасным вариантом). */
export function imageMime(ext: string): string {
  const e = ext.toLowerCase().replace(/^\./, '');
  return IMAGE_MIME[e] || 'image/' + (e || 'png');
}

/** Проверка: относится ли расширение к изображениям. */
export function isImageExtension(ext: string): boolean {
  const e = ext.toLowerCase().replace(/^\./, '');
  return (IMAGE_EXTENSIONS as readonly string[]).includes(e);
}

/** Данные изображения в виде data-URL для <img> и canvas. */
export function imageDataUrl(base64: string, ext: string): string {
  return 'data:' + imageMime(ext) + ';base64,' + base64;
}
