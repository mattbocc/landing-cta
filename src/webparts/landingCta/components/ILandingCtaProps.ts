export type HeroHeightPreset = 'viewport' | 'tall' | 'medium' | 'compact';
export type HeroTextAlignment = 'center' | 'left';
export type FontWeightOption = '400' | '500' | '600' | '700' | '800';

export interface ILandingCtaProps {
  heading: string;
  subheading: string;
  buttonText: string;
  buttonUrl: string;
  openInNewTab: boolean;
  backgroundImage: string;
  backgroundColor: string;
  overlayOpacity: number;
  heightPreset: HeroHeightPreset;
  textAlignment: HeroTextAlignment;
  headingSize: number;
  headingWeight: FontWeightOption;
  headingItalic: boolean;
  headingColor: string;
  subheadingSize: number;
  subheadingWeight: FontWeightOption;
  subheadingItalic: boolean;
  subheadingColor: string;
  buttonTextSize: number;
  buttonWeight: FontWeightOption;
  buttonTextColor: string;
  buttonBackgroundColor: string;
  buttonPaddingY: number;
  buttonPaddingX: number;
  buttonBorderRadius: number;
}
