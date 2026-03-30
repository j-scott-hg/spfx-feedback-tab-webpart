import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CustomFeedbackWebPartStrings';
import CustomFeedback from './components/CustomFeedback';
import { ICustomFeedbackProps, IFeedbackWebPartProps, TabStyle } from './components/IFeedbackTypes';
import { FeedbackService } from './services/FeedbackService';

export interface ICustomFeedbackWebPartProps extends IFeedbackWebPartProps {
  // intentionally empty — all props come from IFeedbackWebPartProps
}

export default class CustomFeedbackWebPart extends BaseClientSideWebPart<ICustomFeedbackWebPartProps> {

  private feedbackService: FeedbackService | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();

    this.feedbackService = new FeedbackService(
      this.context,
      this.context.pageContext.web.absoluteUrl
    );

    // Silently ensure the list exists on first load; do not block render
    this.feedbackService.ensureList().catch((err) => {
      console.warn('[CustomFeedback] Could not ensure Feedback Responses list:', err);
    });
  }

  public render(): void {
    const props: ICustomFeedbackProps = {
      context: this.context,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      pageUrl: this.context.pageContext.legacyPageContext?.serverRequestPath
        ? `${this.context.pageContext.web.absoluteUrl}${this.context.pageContext.legacyPageContext.serverRequestPath}`
        : window.location.href,
      pageTitle: this.context.pageContext.web.title || document.title,

      tabLabel: this.properties.tabLabel || 'Feedback',
      tabStyle: (this.properties.tabStyle as TabStyle) || 'filled',
      tabFontSize: parseInt(`${this.properties.tabFontSize}`, 10) || 15,
      tabWidth: this.properties.tabWidth || 120,
      tabHeight: this.properties.tabHeight || 48,
      formTitle: this.properties.formTitle || 'Give feedback about this site!',
      formDescription: this.properties.formDescription || 'Your feedback is useful to keep improving the site experience!',
      showDescription: this.properties.showDescription !== false,
      questionLabel: this.properties.questionLabel || 'What do you think about this section?',
      showQuestionLabel: this.properties.showQuestionLabel !== false,
      commentPlaceholder: this.properties.commentPlaceholder || 'Leave your comment here',
      submitButtonText: this.properties.submitButtonText || 'Send feedback',
      successMessage: this.properties.successMessage || 'Thank you! Your feedback has been submitted.',
      showStarRating: this.properties.showStarRating !== false
    };

    const element: React.ReactElement<ICustomFeedbackProps> = React.createElement(
      CustomFeedback,
      props
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
              groupName: strings.FeedbackGroupName,
              groupFields: [
                PropertyPaneTextField('tabLabel', {
                  label: strings.TabLabelFieldLabel,
                  placeholder: 'Feedback',
                  description: 'Text shown on the fixed tab in the lower-right corner.'
                }),
                PropertyPaneDropdown('tabStyle', {
                  label: strings.TabStyleFieldLabel,
                  options: [
                    { key: 'filled', text: 'Filled (blue background, white text)' },
                    { key: 'outlined', text: 'Outlined (white background, blue border)' }
                  ],
                  selectedKey: this.properties.tabStyle || 'filled'
                }),
                PropertyPaneTextField('tabFontSize', {
                  label: strings.TabFontSizeFieldLabel,
                  placeholder: '15',
                  description: 'Font size in pixels (e.g. 14). Default: 15.'
                }),
                PropertyPaneSlider('tabWidth', {
                  label: strings.TabWidthFieldLabel,
                  min: 80,
                  max: 400,
                  step: 1,
                  showValue: true,
                  value: this.properties.tabWidth || 120
                }),
                PropertyPaneSlider('tabHeight', {
                  label: strings.TabHeightFieldLabel,
                  min: 32,
                  max: 80,
                  step: 1,
                  showValue: true,
                  value: this.properties.tabHeight || 48
                }),
                PropertyPaneTextField('formTitle', {
                  label: strings.FormTitleFieldLabel,
                  placeholder: 'Give feedback about this site!'
                }),
                PropertyPaneToggle('showDescription', {
                  label: strings.ShowDescriptionFieldLabel,
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneTextField('formDescription', {
                  label: strings.FormDescriptionFieldLabel,
                  placeholder: 'Your feedback is useful to keep improving the site experience!',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneToggle('showQuestionLabel', {
                  label: strings.ShowQuestionLabelFieldLabel,
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneTextField('questionLabel', {
                  label: strings.QuestionLabelFieldLabel,
                  placeholder: 'What do you think about this section?'
                }),
                PropertyPaneTextField('commentPlaceholder', {
                  label: strings.CommentPlaceholderFieldLabel,
                  placeholder: 'Leave your comment here'
                }),
                PropertyPaneToggle('showStarRating', {
                  label: strings.ShowStarRatingFieldLabel,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneTextField('submitButtonText', {
                  label: strings.SubmitButtonTextFieldLabel,
                  placeholder: 'Send feedback'
                }),
                PropertyPaneTextField('successMessage', {
                  label: strings.SuccessMessageFieldLabel,
                  placeholder: 'Thank you! Your feedback has been submitted.',
                  multiline: true,
                  rows: 2
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
