import * as React from 'react';
import styles from './SpfxIssueDetails.module.scss';
import type { ISpfxIssueDetailsProps } from './ISpfxIssueDetailsProps';

export default class SpfxIssueDetails extends React.Component<ISpfxIssueDetailsProps> {
  public render(): React.ReactElement<ISpfxIssueDetailsProps> {
    return (
      <section className={styles.spfxIssueDetails}></section>
    );
  }
}
