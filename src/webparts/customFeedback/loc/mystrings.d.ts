declare interface ICustomFeedbackWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;

  // Feedback-specific property pane strings
  FeedbackGroupName: string;
  TabLabelFieldLabel: string;
  TabStyleFieldLabel: string;
  TabFontSizeFieldLabel: string;
  TabWidthFieldLabel: string;
  TabHeightFieldLabel: string;
  FormTitleFieldLabel: string;
  FormDescriptionFieldLabel: string;
  ShowDescriptionFieldLabel: string;
  QuestionLabelFieldLabel: string;
  ShowQuestionLabelFieldLabel: string;
  CommentPlaceholderFieldLabel: string;
  SubmitButtonTextFieldLabel: string;
  SuccessMessageFieldLabel: string;
  ShowStarRatingFieldLabel: string;
}

declare module 'CustomFeedbackWebPartStrings' {
  const strings: ICustomFeedbackWebPartStrings;
  export = strings;
}
