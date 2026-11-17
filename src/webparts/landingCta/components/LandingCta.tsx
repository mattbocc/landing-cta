import * as React from 'react';
import styles from './LandingCta.module.scss';
import type { HeroHeightPreset, ILandingCtaProps } from './ILandingCtaProps';
import { sanitizeCssColor, sanitizeUrl, toCssBackgroundUrl } from '../utils/urlUtils';

const HEIGHT_CLASS_MAP: Record<HeroHeightPreset, string> = {
  viewport: styles.heightViewport,
  tall: styles.heightTall,
  medium: styles.heightMedium,
  compact: styles.heightCompact
};

const DEFAULT_BG_COLOR = '#163b58';
const DEFAULT_HEADING_COLOR = '#ffffff';
const DEFAULT_SUBHEADING_COLOR = '#f0f4f8';
const DEFAULT_BUTTON_TEXT_COLOR = '#163b58';
const DEFAULT_BUTTON_BG_COLOR = '#ffffff';

function clampNumber(value: number, min: number, max: number, fallback: number): number {
  if (!Number.isFinite(value)) {
    return fallback;
  }
  return Math.max(min, Math.min(max, value));
}

const LandingCta: React.FC<ILandingCtaProps> = (props) => {
  const {
    heading,
    subheading,
    buttonText,
    buttonUrl,
    openInNewTab,
    backgroundImage,
    backgroundColor,
    overlayOpacity,
    heightPreset,
    textAlignment,
    headingSize,
    headingWeight,
    headingItalic,
    headingColor,
    subheadingSize,
    subheadingWeight,
    subheadingItalic,
    subheadingColor,
    buttonTextSize,
    buttonWeight,
    buttonTextColor,
    buttonBackgroundColor,
    buttonPaddingY,
    buttonPaddingX,
    buttonBorderRadius
  } = props;

  const trimmedHeading = (heading || '').trim();
  const trimmedSubheading = (subheading || '').trim();
  const trimmedButtonText = (buttonText || '').trim();
  const safeButtonUrl = sanitizeUrl(buttonUrl);
  const cssBackgroundImage = toCssBackgroundUrl(backgroundImage);
  const safeBackgroundColor = sanitizeCssColor(backgroundColor, DEFAULT_BG_COLOR);

  const clampedOpacity = clampNumber(overlayOpacity, 0, 90, 45) / 100;
  const heightClass = HEIGHT_CLASS_MAP[heightPreset] || styles.heightViewport;
  const alignClass = textAlignment === 'left' ? styles.alignLeft : styles.alignCenter;

  const safeHeadingSize = clampNumber(headingSize, 24, 84, 52);
  const safeSubheadingSize = clampNumber(subheadingSize, 14, 36, 20);
  const safeButtonTextSize = clampNumber(buttonTextSize, 12, 26, 16);
  const safeButtonPadY = clampNumber(buttonPaddingY, 4, 28, 10);
  const safeButtonPadX = clampNumber(buttonPaddingX, 10, 56, 24);
  const safeButtonRadius = clampNumber(buttonBorderRadius, 0, 40, 6);

  const sectionStyle: React.CSSProperties = {
    backgroundColor: safeBackgroundColor,
    ...(cssBackgroundImage ? { backgroundImage: cssBackgroundImage } : {})
  };

  const headingStyle: React.CSSProperties = {
    fontSize: `clamp(${Math.round(safeHeadingSize * 0.62)}px, 4.5vw, ${safeHeadingSize}px)`,
    fontWeight: Number(headingWeight) || 700,
    fontStyle: headingItalic ? 'italic' : 'normal',
    color: sanitizeCssColor(headingColor, DEFAULT_HEADING_COLOR)
  };

  const subheadingStyle: React.CSSProperties = {
    fontSize: `clamp(${Math.round(safeSubheadingSize * 0.85)}px, 1.8vw, ${safeSubheadingSize}px)`,
    fontWeight: Number(subheadingWeight) || 400,
    fontStyle: subheadingItalic ? 'italic' : 'normal',
    color: sanitizeCssColor(subheadingColor, DEFAULT_SUBHEADING_COLOR)
  };

  const buttonStyle: React.CSSProperties = {
    fontSize: `${safeButtonTextSize}px`,
    fontWeight: Number(buttonWeight) || 600,
    color: sanitizeCssColor(buttonTextColor, DEFAULT_BUTTON_TEXT_COLOR),
    backgroundColor: sanitizeCssColor(buttonBackgroundColor, DEFAULT_BUTTON_BG_COLOR),
    padding: `${safeButtonPadY}px ${safeButtonPadX}px`,
    borderRadius: `${safeButtonRadius}px`
  };

  return (
    <section
      className={`${styles.landingCta} ${heightClass} ${alignClass}`}
      style={sectionStyle}
      aria-label={trimmedHeading || 'Landing hero'}
    >
      <div
        className={styles.overlay}
        style={{ backgroundColor: `rgba(8, 18, 30, ${clampedOpacity})` }}
        aria-hidden="true"
      />
      <div className={styles.content}>
        {trimmedHeading && (
          <h1 className={styles.heading} style={headingStyle}>
            {trimmedHeading}
          </h1>
        )}
        {trimmedSubheading && (
          <p className={styles.subheading} style={subheadingStyle}>
            {trimmedSubheading}
          </p>
        )}
        {trimmedButtonText && safeButtonUrl && (
          <div className={styles.ctaWrapper}>
            <a
              className={styles.ctaButton}
              style={buttonStyle}
              href={safeButtonUrl}
              target={openInNewTab ? '_blank' : undefined}
              rel={openInNewTab ? 'noopener noreferrer' : undefined}
            >
              {trimmedButtonText}
            </a>
          </div>
        )}
      </div>
    </section>
  );
};

export default LandingCta;
