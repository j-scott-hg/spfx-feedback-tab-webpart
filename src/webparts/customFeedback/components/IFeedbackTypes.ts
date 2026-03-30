import { WebPartContext } from '@microsoft/sp-webpart-base';

export type TabStyle = 'filled' | 'outlined';

export interface IFeedbackWebPartProps {
  tabLabel: string;
  tabStyle: TabStyle;
  tabFontSize: number;
  tabWidth: number;
  tabHeight: number;
  formTitle: string;
  formDescription: string;
  showDescription: boolean;
  questionLabel: string;
  showQuestionLabel: boolean;
  commentPlaceholder: string;
  submitButtonText: string;
  successMessage: string;
  showStarRating: boolean;
}

export interface ICustomFeedbackProps extends IFeedbackWebPartProps {
  context: WebPartContext;
  siteUrl: string;
  pageUrl: string;
  pageTitle: string;
}

export interface IFeedbackPanelProps extends IFeedbackWebPartProps {
  isOpen: boolean;
  onDismiss: () => void;
  context: WebPartContext;
  siteUrl: string;
  pageUrl: string;
  pageTitle: string;
}

export interface IFeedbackFormProps extends IFeedbackWebPartProps {
  context: WebPartContext;
  siteUrl: string;
  pageUrl: string;
  pageTitle: string;
  onSubmitSuccess: () => void;
}

export interface IStarRatingProps {
  value: number;
  onChange: (rating: number) => void;
  disabled?: boolean;
}

export interface IFeedbackSubmission {
  Title: string;
  CommentText: string;
  Rating?: number;
  PageURL: string;
  PageTitle: string;
  SiteURL: string;
  SubmittedDateTime: string;
}
