import {
  sanitizeCssColor,
  sanitizeUrl,
  toCssBackgroundUrl,
  validateHexColorInput,
  validateUrlInput
} from './urlUtils';

describe('urlUtils', () => {
  describe('sanitizeUrl', () => {
    it('allows https URLs', () => {
      expect(sanitizeUrl('https://outlook.office365.com/book/bookings@example.com/')).toBe(
        'https://outlook.office365.com/book/bookings@example.com/'
      );
    });

    it('allows single-slash site-relative SharePoint paths', () => {
      expect(sanitizeUrl('/sites/bvai/SiteAssets/hero.jpg')).toBe('/sites/bvai/SiteAssets/hero.jpg');
    });

    it('rejects javascript:, data:, and http: URLs', () => {
      const jsUrl = ['javascript', 'alert(1)'].join(':');
      expect(sanitizeUrl(jsUrl)).toBeUndefined();
      expect(sanitizeUrl(`  ${jsUrl.toUpperCase()}  `)).toBeUndefined();
      expect(sanitizeUrl('data:text/html,<script>alert(1)</script>')).toBeUndefined();
      expect(sanitizeUrl('http://example.com')).toBeUndefined();
    });

    it('rejects protocol-relative URLs and control characters', () => {
      expect(sanitizeUrl('//evil.example.com/phish')).toBeUndefined();
      expect(sanitizeUrl('/\\evil.example.com')).toBeUndefined();
      expect(sanitizeUrl('https://example.com/\nfoo')).toBeUndefined();
    });

    it('returns undefined for empty input', () => {
      expect(sanitizeUrl('')).toBeUndefined();
      expect(sanitizeUrl('   ')).toBeUndefined();
      expect(sanitizeUrl(undefined)).toBeUndefined();
    });
  });

  describe('validateUrlInput', () => {
    it('returns empty string for empty or valid URLs', () => {
      expect(validateUrlInput('')).toBe('');
      expect(validateUrlInput('https://example.com/book')).toBe('');
      expect(validateUrlInput('/sites/bvai/SitePages/Home.aspx')).toBe('');
    });

    it('returns an error message for unsafe or invalid URLs', () => {
      const jsUrl = ['javascript', 'alert(1)'].join(':');
      expect(validateUrlInput(jsUrl)).not.toBe('');
      expect(validateUrlInput('not-a-url')).not.toBe('');
    });
  });

  describe('sanitizeCssColor & validateHexColorInput', () => {
    it('accepts valid 3, 6, and 8 digit hex colors', () => {
      expect(sanitizeCssColor('#fff', '#163b58')).toBe('#fff');
      expect(sanitizeCssColor(' #163b58 ', '#000000')).toBe('#163b58');
      expect(validateHexColorInput('#163b58')).toBe('');
      expect(validateHexColorInput('')).toBe('');
    });

    it('falls back and reports error on invalid colors', () => {
      expect(sanitizeCssColor('red; background: url(x)', '#163b58')).toBe('#163b58');
      expect(validateHexColorInput('blue')).not.toBe('');
    });
  });

  describe('toCssBackgroundUrl', () => {
    it('formats safe URLs and rejects unsafe ones', () => {
      const jsUrl = ['javascript', 'alert(1)'].join(':');
      expect(toCssBackgroundUrl('https://cdn.example.com/hero.jpg')).toBe(
        'url("https://cdn.example.com/hero.jpg")'
      );
      expect(toCssBackgroundUrl(jsUrl)).toBeUndefined();
    });
  });
});
