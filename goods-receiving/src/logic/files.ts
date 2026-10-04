/**
 * File name and file type helpers for attachments.
 */

export function getExtension(fileName: string): string {
  const index = fileName.lastIndexOf('.');
  return index > 0 ? fileName.substring(index).toLowerCase() : '';
}

function removeControlCharacters(text: string): string {
  let result = '';
  for (let i = 0; i < text.length; i++) {
    if (text.charCodeAt(i) >= 32) result += text.charAt(i);
  }
  return result;
}

/**
 * Makes a file name safe for SharePoint: removes characters SharePoint does not
 * allow (" * : < > ? / \ |), leading/trailing spaces and dots, and the "~$" prefix.
 */
export function sanitizeFileName(fileName: string): string {
  let name = removeControlCharacters((fileName || '').replace(/["*:<>?/\\|]/g, '_'));
  name = name.replace(/^~\$/, '').trim().replace(/^\.+/, '').replace(/[.\s]+$/, '');
  if (!name) name = 'file';
  // SharePoint limits the full path; keep names reasonably short.
  if (name.length > 120) {
    const extension = getExtension(name);
    name = name.substring(0, 120 - extension.length) + extension;
  }
  return name;
}

/**
 * Returns a name that is not in `existingNames` (case-insensitive) by adding
 * " (2)", " (3)", ... before the extension. Phones name every photo "image.jpg".
 */
export function makeUniqueFileName(fileName: string, existingNames: string[]): string {
  const taken: Record<string, boolean> = {};
  existingNames.forEach((n) => {
    taken[n.toLowerCase()] = true;
  });
  if (!taken[fileName.toLowerCase()]) return fileName;
  const extension = getExtension(fileName);
  const base = extension ? fileName.substring(0, fileName.length - extension.length) : fileName;
  for (let i = 2; i < 1000; i++) {
    const candidate = `${base} (${i})${extension}`;
    if (!taken[candidate.toLowerCase()]) return candidate;
  }
  return `${base} (${new Date().getTime()})${extension}`;
}

/** Validates type and size. Returns an error message or undefined. */
export function validateFile(file: { name: string; size: number }, allowedExtensions: string[], maxFileSizeMB: number): string | undefined {
  const extension = getExtension(file.name);
  if (allowedExtensions.indexOf(extension) < 0) {
    return `"${file.name}" cannot be uploaded. Allowed file types: ${allowedExtensions.join(', ')}.`;
  }
  if (file.size === 0) {
    return `"${file.name}" is empty.`;
  }
  if (file.size > maxFileSizeMB * 1024 * 1024) {
    return `"${file.name}" is larger than ${maxFileSizeMB} MB.`;
  }
  return undefined;
}

/** Image types every browser can display inline as a thumbnail. HEIC is not one of them. */
export function isBrowserImage(fileName: string): boolean {
  return ['.jpg', '.jpeg', '.png', '.gif', '.webp'].indexOf(getExtension(fileName)) >= 0;
}

/** Fluent UI icon name for a file. */
export function getFileIconName(fileName: string): string {
  switch (getExtension(fileName)) {
    case '.pdf':
      return 'PDF';
    case '.jpg':
    case '.jpeg':
    case '.png':
    case '.heic':
    case '.heif':
      return 'FileImage';
    case '.xlsx':
      return 'ExcelDocument';
    case '.docx':
      return 'WordDocument';
    case '.msg':
      return 'Mail';
    default:
      return 'Page';
  }
}
