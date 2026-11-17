import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneCustomFieldProps,
  type IPropertyPaneDropdownOption,
  type IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneFieldType,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'LandingCtaWebPartStrings';
import LandingCta from './components/LandingCta';
import { ImagePickerField } from './components/ImagePickerField';
import type {
  FontWeightOption,
  HeroHeightPreset,
  HeroTextAlignment,
  ILandingCtaProps
} from './components/ILandingCtaProps';
import { validateHexColorInput, validateUrlInput } from './utils/urlUtils';

export interface ILandingCtaWebPartProps {
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

const FONT_WEIGHT_OPTIONS: IPropertyPaneDropdownOption[] = [
  { key: '400', text: 'Regular (400)' },
  { key: '500', text: 'Medium (500)' },
  { key: '600', text: 'Semi-Bold (600)' },
  { key: '700', text: 'Bold (700)' },
  { key: '800', text: 'Extra-Bold (800)' }
];

export default class LandingCtaWebPart extends BaseClientSideWebPart<ILandingCtaWebPartProps> {

  private _root: Root | undefined;
  private _imagePickerRoot: Root | undefined;

  public render(): void {
    const element: React.ReactElement<ILandingCtaProps> = React.createElement(
      LandingCta,
      {
        heading: this.properties.heading ?? 'Welcome to Blue Vest AI',
        subheading: this.properties.subheading ?? 'Explore our services and connect with our team.',
        buttonText: this.properties.buttonText ?? 'Book an Appointment',
        buttonUrl: this.properties.buttonUrl ?? '',
        openInNewTab: this.properties.openInNewTab ?? true,
        backgroundImage: this.properties.backgroundImage ?? '',
        backgroundColor: this.properties.backgroundColor ?? '#163b58',
        overlayOpacity: this.properties.overlayOpacity ?? 45,
        heightPreset: this.properties.heightPreset ?? 'viewport',
        textAlignment: this.properties.textAlignment ?? 'center',
        headingSize: this.properties.headingSize ?? 52,
        headingWeight: this.properties.headingWeight ?? '700',
        headingItalic: this.properties.headingItalic ?? false,
        headingColor: this.properties.headingColor ?? '#ffffff',
        subheadingSize: this.properties.subheadingSize ?? 20,
        subheadingWeight: this.properties.subheadingWeight ?? '400',
        subheadingItalic: this.properties.subheadingItalic ?? false,
        subheadingColor: this.properties.subheadingColor ?? '#f0f4f8',
        buttonTextSize: this.properties.buttonTextSize ?? 16,
        buttonWeight: this.properties.buttonWeight ?? '600',
        buttonTextColor: this.properties.buttonTextColor ?? '#163b58',
        buttonBackgroundColor: this.properties.buttonBackgroundColor ?? '#ffffff',
        buttonPaddingY: this.properties.buttonPaddingY ?? 10,
        buttonPaddingX: this.properties.buttonPaddingX ?? 24,
        buttonBorderRadius: this.properties.buttonBorderRadius ?? 6
      }
    );

    if (!this._root) {
      this._root = createRoot(this.domElement);
    }

    this._root.render(element);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    this._imagePickerRoot?.unmount();
    this._imagePickerRoot = undefined;
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.ContentGroupName,
              groupFields: [
                PropertyPaneTextField('heading', {
                  label: strings.HeadingFieldLabel
                }),
                PropertyPaneTextField('subheading', {
                  label: strings.SubheadingFieldLabel,
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneTextField('buttonText', {
                  label: strings.ButtonTextFieldLabel
                }),
                PropertyPaneTextField('buttonUrl', {
                  label: strings.ButtonUrlFieldLabel,
                  placeholder: 'https://outlook.office365.com/book/...',
                  onGetErrorMessage: validateUrlInput,
                  deferredValidationTime: 300
                }),
                PropertyPaneToggle('openInNewTab', {
                  label: strings.OpenInNewTabFieldLabel
                })
              ]
            },
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                {
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: 'backgroundImage',
                  properties: {
                    key: 'backgroundImagePicker',
                    onRender: (domElement: HTMLElement) => {
                      if (!this._imagePickerRoot) {
                        this._imagePickerRoot = createRoot(domElement);
                      }
                      this._imagePickerRoot.render(
                        React.createElement(ImagePickerField, {
                          label: strings.BackgroundImageFieldLabel,
                          value: this.properties.backgroundImage ?? '',
                          context: this.context,
                          onChange: (newUrl: string) => {
                            this.properties.backgroundImage = newUrl;
                            this.render();
                            this.context.propertyPane.refresh();
                          }
                        })
                      );
                    },
                    onDispose: () => {
                      this._imagePickerRoot?.unmount();
                      this._imagePickerRoot = undefined;
                    }
                  }
                } as IPropertyPaneField<IPropertyPaneCustomFieldProps>,
                PropertyPaneTextField('backgroundColor', {
                  label: strings.BackgroundColorFieldLabel,
                  placeholder: '#163b58',
                  onGetErrorMessage: validateHexColorInput,
                  deferredValidationTime: 300
                }),
                PropertyPaneSlider('overlayOpacity', {
                  label: strings.OverlayOpacityFieldLabel,
                  min: 0,
                  max: 90,
                  step: 5,
                  showValue: true
                }),
                PropertyPaneDropdown('heightPreset', {
                  label: strings.HeightPresetFieldLabel,
                  options: [
                    { key: 'viewport', text: 'Full viewport (below SharePoint header)' },
                    { key: 'tall', text: 'Tall (620px)' },
                    { key: 'medium', text: 'Medium (480px)' },
                    { key: 'compact', text: 'Compact (340px)' }
                  ]
                }),
                PropertyPaneDropdown('textAlignment', {
                  label: strings.TextAlignmentFieldLabel,
                  options: [
                    { key: 'center', text: 'Centered' },
                    { key: 'left', text: 'Left aligned' }
                  ]
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.StylePageDescription
          },
          groups: [
            {
              groupName: strings.HeadingStyleGroupName,
              groupFields: [
                PropertyPaneSlider('headingSize', {
                  label: strings.HeadingSizeFieldLabel,
                  min: 24,
                  max: 84,
                  step: 2,
                  showValue: true
                }),
                PropertyPaneDropdown('headingWeight', {
                  label: strings.HeadingWeightFieldLabel,
                  options: FONT_WEIGHT_OPTIONS
                }),
                PropertyPaneToggle('headingItalic', {
                  label: strings.HeadingItalicFieldLabel
                }),
                PropertyPaneTextField('headingColor', {
                  label: strings.HeadingColorFieldLabel,
                  placeholder: '#ffffff',
                  onGetErrorMessage: validateHexColorInput,
                  deferredValidationTime: 300
                })
              ]
            },
            {
              groupName: strings.SubheadingStyleGroupName,
              groupFields: [
                PropertyPaneSlider('subheadingSize', {
                  label: strings.SubheadingSizeFieldLabel,
                  min: 14,
                  max: 36,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneDropdown('subheadingWeight', {
                  label: strings.SubheadingWeightFieldLabel,
                  options: FONT_WEIGHT_OPTIONS
                }),
                PropertyPaneToggle('subheadingItalic', {
                  label: strings.SubheadingItalicFieldLabel
                }),
                PropertyPaneTextField('subheadingColor', {
                  label: strings.SubheadingColorFieldLabel,
                  placeholder: '#f0f4f8',
                  onGetErrorMessage: validateHexColorInput,
                  deferredValidationTime: 300
                })
              ]
            },
            {
              groupName: strings.ButtonStyleGroupName,
              groupFields: [
                PropertyPaneSlider('buttonTextSize', {
                  label: strings.ButtonTextSizeFieldLabel,
                  min: 12,
                  max: 26,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneDropdown('buttonWeight', {
                  label: strings.ButtonWeightFieldLabel,
                  options: FONT_WEIGHT_OPTIONS
                }),
                PropertyPaneTextField('buttonTextColor', {
                  label: strings.ButtonTextColorFieldLabel,
                  placeholder: '#163b58',
                  onGetErrorMessage: validateHexColorInput,
                  deferredValidationTime: 300
                }),
                PropertyPaneTextField('buttonBackgroundColor', {
                  label: strings.ButtonBackgroundColorFieldLabel,
                  placeholder: '#ffffff',
                  onGetErrorMessage: validateHexColorInput,
                  deferredValidationTime: 300
                }),
                PropertyPaneSlider('buttonPaddingY', {
                  label: strings.ButtonPaddingYFieldLabel,
                  min: 4,
                  max: 28,
                  step: 2,
                  showValue: true
                }),
                PropertyPaneSlider('buttonPaddingX', {
                  label: strings.ButtonPaddingXFieldLabel,
                  min: 10,
                  max: 56,
                  step: 2,
                  showValue: true
                }),
                PropertyPaneSlider('buttonBorderRadius', {
                  label: strings.ButtonBorderRadiusFieldLabel,
                  min: 0,
                  max: 40,
                  step: 2,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
