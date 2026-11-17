const HEX_COLOR_REGEX: RegExp = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

/**
 * Sanitizes a user-supplied URL for navigation or background images.
 * Allows https:// URLs and single-slash site-relative paths (e.g. /sites/...).
 * Rejects javascript:, data:, http:, protocol-relative (//), and malformed URLs.
 */
export function sanitizeUrl(rawUrl: string | undefined): string | undefined {
  if (!rawUrl) {
    return undefined;
  }

  const trimmed = rawUrl.trim();
  if (!trimmed) {
    return undefined;
  }

  // Reject ASCII control characters or newlines that could obfuscate schemes
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001F\u007F]/.test(trimmed)) {
    return undefined;
  }

  // Allow single-slash site-relative paths (e.g., /sites/bvai/SiteAssets/hero.jpg)
  // Reject protocol-relative // and backslash variants /\
  if (trimmed.startsWith('/')) {
    if (trimmed.startsWith('//') || trimmed.startsWith('/\\')) {
      return undefined;
    }
    return trimmed;
  }

  try {
    const parsed = new URL(trimmed);
    if (parsed.protocol === 'https:') {
      return parsed.href;
    }
    return undefined;
  } catch {
    return undefined;
  }
}

/**
 * Property pane validator for optional URL fields.
 * Returns an empty string when valid (or empty), or a user-facing error message.
 */
export function validateUrlInput(rawUrl: string): string {
  const trimmed = (rawUrl || '').trim();
  if (!trimmed) {
    return '';
  }

  if (sanitizeUrl(trimmed) !== undefined) {
    return '';
  }

  return 'Enter a valid https:// URL or a site-relative path starting with /.';
}

/**
 * Sanitizes a CSS hex color string, returning the fallback if invalid.
 */
export function sanitizeCssColor(rawColor: string | undefined, fallback: string): string {
  const trimmed = (rawColor || '').trim();
  if (HEX_COLOR_REGEX.test(trimmed)) {
    return trimmed;
  }
  return fallback;
}

/**
 * Property pane validator for optional hex color fields.
 */
export function validateHexColorInput(rawColor: string): string {
  const trimmed = (rawColor || '').trim();
  if (!trimmed || HEX_COLOR_REGEX.test(trimmed)) {
    return '';
  }
  return 'Enter a valid hex color (for example, #163b58).';
}

/**
 * Wraps a sanitized URL in a CSS url("...") value with quote/backslash escaping.
 */
export function toCssBackgroundUrl(rawUrl: string | undefined): string | undefined {
  const safeUrl = sanitizeUrl(rawUrl);
  if (!safeUrl) {
    return undefined;
  }

  const escaped = safeUrl.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  return `url("${escaped}")`;
}

