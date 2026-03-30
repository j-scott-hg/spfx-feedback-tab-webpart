import * as React from 'react';
import { ICustomFeedbackProps } from './IFeedbackTypes';
import FeedbackPanel from './FeedbackPanel';
import styles from './FeedbackPanel.module.scss';

interface ICustomFeedbackState {
  isPanelOpen: boolean;
}

export default class CustomFeedback extends React.Component<ICustomFeedbackProps, ICustomFeedbackState> {
  constructor(props: ICustomFeedbackProps) {
    super(props);
    this.state = { isPanelOpen: false };
  }

  private openPanel = (): void => {
    this.setState({ isPanelOpen: true });
  };

  private closePanel = (): void => {
    this.setState({ isPanelOpen: false });
  };

  public render(): React.ReactElement {
    const { tabLabel, tabStyle, tabFontSize, tabWidth, tabHeight, ...rest } = this.props;
    const { isPanelOpen } = this.state;

    const styleClass = tabStyle === 'outlined'
      ? styles.feedbackTabOutlined
      : styles.feedbackTabFilled;

    const tabInlineStyle: React.CSSProperties = {
      fontSize: `${tabFontSize || 15}px`,
      width: `${tabWidth || 120}px`,
      height: `${tabHeight || 48}px`
    };

    return (
      <>
        {/* Fixed feedback tab — position is fixed to viewport, not the web part */}
        <button
          className={`${styles.feedbackTab} ${styleClass}`}
          style={tabInlineStyle}
          onClick={this.openPanel}
          aria-haspopup="dialog"
          aria-expanded={isPanelOpen}
          aria-label={tabLabel || 'Feedback'}
        >
          {tabLabel || 'Feedback'}
        </button>

        <FeedbackPanel
          {...rest}
          tabLabel={tabLabel}
          tabStyle={tabStyle}
          tabFontSize={tabFontSize}
          tabWidth={tabWidth}
          tabHeight={tabHeight}
          isOpen={isPanelOpen}
          onDismiss={this.closePanel}
        />
      </>
    );
  }
}
