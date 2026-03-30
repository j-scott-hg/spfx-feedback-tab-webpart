import * as React from 'react';
import { IconButton } from '@fluentui/react';
import { IFeedbackPanelProps } from './IFeedbackTypes';
import FeedbackForm from './FeedbackForm';
import styles from './FeedbackPanel.module.scss';

interface IFeedbackPanelState {
  submitted: boolean;
}

export default class FeedbackPanel extends React.Component<IFeedbackPanelProps, IFeedbackPanelState> {
  private panelRef = React.createRef<HTMLDivElement>();

  constructor(props: IFeedbackPanelProps) {
    super(props);
    this.state = { submitted: false };
  }

  public componentDidUpdate(prevProps: IFeedbackPanelProps): void {
    if (!prevProps.isOpen && this.props.isOpen) {
      this.setState({ submitted: false });
      setTimeout(() => {
        if (this.panelRef.current) {
          this.panelRef.current.focus();
        }
      }, 100);
    }
  }

  private handleKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Escape') {
      this.props.onDismiss();
    }
  };

  private handleSubmitSuccess = (): void => {
    this.setState({ submitted: true });
  };

  public render(): React.ReactElement {
    const { isOpen, onDismiss } = this.props;

    return (
      <>
        {/* Backdrop overlay */}
        {isOpen && (
          <div
            className={styles.backdrop}
            onClick={onDismiss}
            aria-hidden="true"
          />
        )}

        {/* Slide-in panel */}
        <div
          ref={this.panelRef}
          className={`${styles.feedbackPanel} ${isOpen ? styles.panelOpen : ''}`}
          role="dialog"
          aria-modal="true"
          aria-label="Feedback panel"
          tabIndex={-1}
          onKeyDown={this.handleKeyDown}
        >
          <div className={styles.panelHeader}>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title="Close feedback panel"
              ariaLabel="Close feedback panel"
              onClick={onDismiss}
              className={styles.closeButton}
            />
          </div>

          <div className={styles.panelContent}>
            <FeedbackForm
              {...this.props}
              onSubmitSuccess={this.handleSubmitSuccess}
            />
          </div>
        </div>
      </>
    );
  }
}
