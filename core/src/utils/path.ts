export function getExtensionFromPath(path: string): string | undefined {
  const normalized = path.trim();
  const lastDotIndex = normalized.lastIndexOf(".");
  if (lastDotIndex === -1 || lastDotIndex === normalized.length - 1) {
    return undefined;
  }

  return normalized.slice(lastDotIndex + 1).toLowerCase();
}

export function getFileName(path: string): string {
  const segments = path.split(/[\\/]/).filter(Boolean);
  return segments.length > 0 ? segments[segments.length - 1] : path;
}

