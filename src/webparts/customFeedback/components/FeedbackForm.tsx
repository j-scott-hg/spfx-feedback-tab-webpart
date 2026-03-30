import * as React from 'react';
import { PrimaryButton, TextField, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
import { IFeedbackFormProps } from './IFeedbackTypes';
import { FeedbackService } from '../services/FeedbackService';
import StarRating from './StarRating';
import styles from './FeedbackPanel.module.scss';

type FormState = 'idle' | 'submitting' | 'success' | 'error';

interface IFeedbackFormState {
  comment: string;
  rating: number;
  formState: FormState;
  errorMessage: string;
  commentError: string;
}

export default class FeedbackForm extends React.Component<IFeedbackFormProps, IFeedbackFormState> {
  private feedbackService: FeedbackService;

  constructor(props: IFeedbackFormProps) {
    super(props);
    this.state = {
      comment: '',
      rating: 0,
      formState: 'idle',
      errorMessage: '',
      commentError: ''
    };
    this.feedbackService = new FeedbackService(props.context, props.siteUrl);
  }

  private handleCommentChange = (_e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({ comment: newValue || '', commentError: '' });
  };

  private handleRatingChange = (rating: number): void => {
    this.setState({ rating });
  };

  private validate(): boolean {
    if (!this.state.comment.trim()) {
      this.setState({ commentError: 'Please enter a comment before submitting.' });
      return false;
    }
    return true;
  }

  private handleSubmit = async (): Promise<void> => {
    if (!this.validate()) return;
    if (this.state.formState === 'submitting') return;

    this.setState({ formState: 'submitting', errorMessage: '' });

    try {
      const now = new Date();
      const title = `Feedback - ${now.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })}`;

      const submission = {
        Title: title,
        CommentText: this.state.comment.trim(),
        ...(this.props.showStarRating && this.state.rating > 0 ? { Rating: this.state.rating } : {}),
        PageURL: this.props.pageUrl,
        PageTitle: this.props.pageTitle,
        SiteURL: this.props.siteUrl,
        SubmittedDateTime: now.toISOString()
      };

      await this.feedbackService.submitFeedback(submission);
      this.setState({ formState: 'success', comment: '', rating: 0 });
      this.props.onSubmitSuccess();
    } catch (err) {
      const message = err instanceof Error ? err.message : 'An unexpected error occurred. Please try again.';
      this.setState({ formState: 'error', errorMessage: message });
    }
  };

  public render(): React.ReactElement {
    const {
      formTitle,
      formDescription,
      showDescription,
      questionLabel,
      showQuestionLabel,
      commentPlaceholder,
      submitButtonText,
      successMessage,
      showStarRating
    } = this.props;

    const { comment, rating, formState, errorMessage, commentError } = this.state;
    const isSubmitting = formState === 'submitting';

    if (formState === 'success') {
      return (
        <div className={styles.successContainer} role="alert" aria-live="polite">
          <div className={styles.successIcon} aria-hidden="true">✓</div>
          <p className={styles.successMessage}>{successMessage || 'Thank you for your feedback!'}</p>
        </div>
      );
    }

    return (
      <div className={styles.feedbackForm}>
        {formTitle && (
          <h2 className={styles.formTitle}>{formTitle}</h2>
        )}

        {showDescription && formDescription && (
          <p className={styles.formDescription}>{formDescription}</p>
        )}

        {formState === 'error' && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={() => this.setState({ formState: 'idle', errorMessage: '' })}
            dismissButtonAriaLabel="Close error"
          >
            {errorMessage}
          </MessageBar>
        )}

        {showQuestionLabel && questionLabel && (
          <p className={styles.questionLabel}>{questionLabel}</p>
        )}

        <TextField
          multiline
          rows={5}
          placeholder={commentPlaceholder || 'Leave your comment here'}
          value={comment}
          onChange={this.handleCommentChange}
          disabled={isSubmitting}
          errorMessage={commentError}
          aria-label={questionLabel || 'Your feedback comment'}
          className={styles.commentField}
          resizable={false}
        />

        {showStarRating && (
          <div className={styles.ratingContainer}>
            <StarRating
              value={rating}
              onChange={this.handleRatingChange}
              disabled={isSubmitting}
            />
          </div>
        )}

        <div className={styles.submitContainer}>
          {isSubmitting ? (
            <Spinner size={SpinnerSize.small} label="Submitting..." ariaLive="assertive" />
          ) : (
            <PrimaryButton
              text={submitButtonText || 'Send feedback'}
              onClick={this.handleSubmit}
              disabled={isSubmitting}
              className={styles.submitButton}
              aria-label={submitButtonText || 'Send feedback'}
            />
          )}
        </div>
      </div>
    );
  }
}
