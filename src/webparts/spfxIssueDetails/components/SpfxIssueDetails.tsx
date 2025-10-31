import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler
} from 'chart.js';
import { Scatter } from 'react-chartjs-2';
import styles from './SpfxIssueDetails.module.scss';
import type { ISpfxIssueDetailsProps } from './ISpfxIssueDetailsProps';

// Register Chart.js components
ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler
);

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

  private _getQuadrantColor = (x: number, y: number): { bg: string; border: string } => {
    // Quadrant 1 (top-right): x >= 25, y >= 25 - Green (high resolvability, high opportunity)
    // Quadrant 2 (top-left): x < 25, y >= 25 - Yellow (low resolvability, high opportunity)
    // Quadrant 3 (bottom-left): x < 25, y < 25 - Red (low resolvability, low opportunity)
    // Quadrant 4 (bottom-right): x >= 25, y < 25 - Orange (high resolvability, low opportunity)
    
    if (x >= 25 && y >= 25) {
      return { bg: 'rgba(75, 192, 192, 0.6)', border: 'rgba(75, 192, 192, 1)' }; // Green - Q1
    } else if (x < 25 && y >= 25) {
      return { bg: 'rgba(255, 206, 86, 0.6)', border: 'rgba(255, 206, 86, 1)' }; // Yellow - Q2
    } else if (x < 25 && y < 25) {
      return { bg: 'rgba(255, 99, 132, 0.6)', border: 'rgba(255, 99, 132, 1)' }; // Red - Q3
    } else {
      return { bg: 'rgba(255, 159, 64, 0.6)', border: 'rgba(255, 159, 64, 1)' }; // Orange - Q4
    }
  }

  private _prepareChartData = () => {
    const { items } = this.state;
    
    // Filter items that have valid Resolvability and Opportunity values
    const chartData = items
      .filter(item => {
        const resolvability = this._parseNumber(item.Resolvability);
        const opportunity = this._parseNumber(item.Opportunity);
        return resolvability !== null && opportunity !== null;
      })
      .map(item => {
        const x = this._parseNumber(item.Resolvability) as number;
        const y = this._parseNumber(item.Opportunity) as number;
        const colors = this._getQuadrantColor(x, y);
        return {
          x: x,
          y: y,
          label: item.Title || `Item ${item.Id}`,
          backgroundColor: colors.bg,
          borderColor: colors.border
        };
      });

    return chartData;
  }

  private _parseNumber = (value: any): number | null => {
    if (value === null || value === undefined || value === '') {
      return null;
    }
    const num = typeof value === 'string' ? parseFloat(value) : Number(value);
    return isNaN(num) ? null : num;
  }

  private _getChartOptions = () => {
    return {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: 1,
      scales: {
        x: {
          type: 'linear' as const,
          position: 'bottom' as const,
          min: 0,
          max: 50,
          title: {
            display: true,
            text: 'Resolvability'
          },
          grid: {
            color: (context: any) => {
              if (context.tick.value === 25) {
                return 'rgba(0, 0, 0, 0.3)'; // Darker line for quadrant divider
              }
              return 'rgba(0, 0, 0, 0.1)';
            },
            lineWidth: (context: any) => {
              if (context.tick.value === 25) {
                return 2; // Thicker line for quadrant divider
              }
              return 1;
            }
          }
        },
        y: {
          type: 'linear' as const,
          min: 0,
          max: 50,
          title: {
            display: true,
            text: 'Opportunity'
          },
          grid: {
            color: (context: any) => {
              if (context.tick.value === 25) {
                return 'rgba(0, 0, 0, 0.3)'; // Darker line for quadrant divider
              }
              return 'rgba(0, 0, 0, 0.1)';
            },
            lineWidth: (context: any) => {
              if (context.tick.value === 25) {
                return 2; // Thicker line for quadrant divider
              }
              return 1;
            }
          }
        }
      },
      plugins: {
        title: {
          display: true,
          text: 'Resolvability vs Opportunity'
        },
        legend: {
          display: false
        },
        tooltip: {
          callbacks: {
            label: (context: any) => {
              const point = context.raw;
              return `${point.label || 'Item'}: (${point.x}, ${point.y})`;
            }
          }
        }
      }
    };
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

    const chartData = this._prepareChartData();

    return (
      <section className={styles.spfxIssueDetails}>
        <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #ddd', marginBottom: '20px' }}>
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
        
        {chartData.length > 0 && (
          <div style={{ marginTop: '20px', padding: '20px', border: '1px solid #ddd', backgroundColor: '#fff' }}>
            <Scatter
              data={{
                datasets: [
                  {
                    label: 'Items',
                    data: chartData.map(point => ({ x: point.x, y: point.y, label: point.label })),
                    pointBackgroundColor: chartData.map(point => point.backgroundColor),
                    pointBorderColor: chartData.map(point => point.borderColor),
                    pointRadius: 6,
                    pointHoverRadius: 8
                  }
                ]
              }}
              options={this._getChartOptions()}
            />
          </div>
        )}
      </section>
    );
  }
}
