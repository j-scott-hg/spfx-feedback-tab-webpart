import * as React from 'react';
import { IStarRatingProps } from './IFeedbackTypes';
import styles from './FeedbackPanel.module.scss';

const StarRating: React.FC<IStarRatingProps> = ({ value, onChange, disabled }) => {
  const [hovered, setHovered] = React.useState<number>(0);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLButtonElement>, index: number): void => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowUp') {
      e.preventDefault();
      onChange(Math.min(5, index + 1));
    } else if (e.key === 'ArrowLeft' || e.key === 'ArrowDown') {
      e.preventDefault();
      onChange(Math.max(0, index - 1));
    }
  };

  return (
    <div
      className={styles.starRating}
      role="radiogroup"
      aria-label="Star rating"
    >
      {[1, 2, 3, 4, 5].map((star) => {
        const filled = hovered > 0 ? star <= hovered : star <= value;
        return (
          <button
            key={star}
            type="button"
            className={`${styles.starButton} ${filled ? styles.starFilled : styles.starEmpty}`}
            onClick={() => !disabled && onChange(star === value ? 0 : star)}
            onMouseEnter={() => !disabled && setHovered(star)}
            onMouseLeave={() => !disabled && setHovered(0)}
            onKeyDown={(e) => !disabled && handleKeyDown(e, star)}
            aria-label={`${star} star${star !== 1 ? 's' : ''}`}
            aria-pressed={star <= value}
            disabled={disabled}
          >
            {filled ? '★' : '☆'}
          </button>
        );
      })}
    </div>
  );
};

export default StarRating;
