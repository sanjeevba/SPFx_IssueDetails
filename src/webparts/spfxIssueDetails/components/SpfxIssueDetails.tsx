import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './SpfxIssueDetails.module.scss';
import type { ISpfxIssueDetailsProps } from './ISpfxIssueDetailsProps';

export interface IListItem {
  Id: number;
  Title?: string;
  Resolvability?: string;
  Opportunity?: string;
  [key: string]: any;
}

export default class SpfxIssueDetails extends React.Component<ISpfxIssueDetailsProps, { items: IListItem[]; loading: boolean; error: string | null }> {
  constructor(props: ISpfxIssueDetailsProps) {
    super(props);
    this.state = {
      items: [],
      loading: true,
      error: null
    };
  }

  public componentDidMount(): void {
    this._fetchListItems();
  }

  public componentDidUpdate(prevProps: ISpfxIssueDetailsProps): void {
    if (prevProps.marketAccessIssueList !== this.props.marketAccessIssueList) {
      this._fetchListItems();
    }
  }

  private _fetchListItems = async (): Promise<void> => {
    const { marketAccessIssueList, context } = this.props;

    if (!marketAccessIssueList) {
      this.setState({ items: [], loading: false, error: 'List name not specified' });
      return;
    }

    this.setState({ loading: true, error: null });

    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const listUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(marketAccessIssueList)}')/items?$select=Id,Title,Resolvability,Opportunity`;

      const response = await context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch list items: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      this.setState({ items: data.value || [], loading: false, error: null });
    } catch (error) {
      this.setState({ items: [], loading: false, error: error.message });
    }
  }

  public render(): React.ReactElement<ISpfxIssueDetailsProps> {
    const { marketAccessIssueList } = this.props;
    const { items, loading, error } = this.state;

    if (loading) {
      return (
        <section className={styles.spfxIssueDetails}>
          <div>Loading list items...</div>
        </section>
      );
    }

    if (error) {
      return (
        <section className={styles.spfxIssueDetails}>
          <div style={{ color: 'red' }}>Error: {error}</div>
        </section>
      );
    }

    return (
      <section className={styles.spfxIssueDetails}>
        <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #ddd' }}>
          <thead>
            <tr style={{ backgroundColor: '#f2f2f2' }}>
              <th style={{ padding: '12px', textAlign: 'left', border: '1px solid #ddd' }}>ID</th>
              <th style={{ padding: '12px', textAlign: 'left', border: '1px solid #ddd' }}>Title</th>
              <th style={{ padding: '12px', textAlign: 'left', border: '1px solid #ddd' }}>Resolvability</th>
              <th style={{ padding: '12px', textAlign: 'left', border: '1px solid #ddd' }}>Opportunity</th>
            </tr>
          </thead>
          <tbody>
            {items.length === 0 ? (
              <tr>
                <td colSpan={4} style={{ padding: '12px', border: '1px solid #ddd', textAlign: 'center' }}>
                  No items found in list "{marketAccessIssueList}"
                </td>
              </tr>
            ) : (
              items.map((item) => (
                <tr key={item.Id}>
                  <td style={{ padding: '12px', border: '1px solid #ddd' }}>{item.Id}</td>
                  <td style={{ padding: '12px', border: '1px solid #ddd' }}>{item.Title || '(No Title)'}</td>
                  <td style={{ padding: '12px', border: '1px solid #ddd' }}>{item.Resolvability || '-'}</td>
                  <td style={{ padding: '12px', border: '1px solid #ddd' }}>{item.Opportunity || '-'}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </section>
    );
  }
}
