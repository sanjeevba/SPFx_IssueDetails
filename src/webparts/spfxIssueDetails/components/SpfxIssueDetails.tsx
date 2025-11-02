import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  Chart as ChartJS,
  type Chart,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler,
} from "chart.js";
import { Scatter } from "react-chartjs-2";
import styles from "./SpfxIssueDetails.module.scss";
import type { ISpfxIssueDetailsProps } from "./ISpfxIssueDetailsProps";

// Define watermark plugin options type
interface WatermarkPluginOptions {
  enabled?: boolean;
  labels?: {
    topRight?: string;
    topLeft?: string;
    lowerRight?: string;
    lowerLeft?: string;
  };
}

// Define watermark plugin
const watermarkPlugin = {
  id: "watermark",
  afterDraw: (chart: Chart) => {
    // Check if watermark should be shown from chart options
    const pluginOptions = (chart.options?.plugins as Record<string, unknown>)
      ?.watermark as WatermarkPluginOptions | undefined;
    const showWatermark = pluginOptions?.enabled !== false;

    if (!showWatermark) return;

    // Get labels from chart options
    const labels = pluginOptions?.labels || {};
    const topRightLabel = labels.topRight || "1 - High Priority";
    const topLeftLabel = labels.topLeft || "2O - Big Impact";
    const lowerRightLabel = labels.lowerRight || "2R - Quick Win";
    const lowerLeftLabel = labels.lowerLeft || "3 - Low Priority";

    const ctx = chart.ctx;
    const chartArea = chart.chartArea;

    if (!chartArea) return; // Guard against undefined chartArea

    // Calculate quadrant centers (each quadrant is 0-25 or 25-50)
    // X-axis (Resolvability): 0-50, divided at 25
    // Y-axis (Opportunity): 0-50, divided at 25

    // Quadrant 1 (top-right): Resolvability >= 25, Opportunity >= 25
    const q1X = chartArea.left + (chartArea.right - chartArea.left) * 0.75; // 37.5 on axis = 75% of chart width
    const q1Y = chartArea.top + (chartArea.bottom - chartArea.top) * 0.25; // 12.5 on axis = 25% from top

    // Quadrant 2 (top-left): Resolvability < 25, Opportunity >= 25
    const q2X = chartArea.left + (chartArea.right - chartArea.left) * 0.25; // 12.5 on axis = 25% of chart width
    const q2Y = chartArea.top + (chartArea.bottom - chartArea.top) * 0.25; // 12.5 on axis = 25% from top

    // Quadrant 3 (bottom-left): Resolvability < 25, Opportunity < 25
    const q3X = chartArea.left + (chartArea.right - chartArea.left) * 0.25; // 12.5 on axis = 25% of chart width
    const q3Y = chartArea.top + (chartArea.bottom - chartArea.top) * 0.75; // 37.5 on axis = 75% from top

    // Quadrant 4 (bottom-right): Resolvability >= 25, Opportunity < 25
    const q4X = chartArea.left + (chartArea.right - chartArea.left) * 0.75; // 37.5 on axis = 75% of chart width
    const q4Y = chartArea.top + (chartArea.bottom - chartArea.top) * 0.75; // 37.5 on axis = 75% from top

    // Save context and set text properties
    ctx.save();
    ctx.font = "bold 14px Arial";
    ctx.fillStyle = "rgba(0, 0, 0, 0.3)";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    // Draw watermarks with dynamic labels
    // Quadrant 1 (top-right)
    ctx.fillText(topRightLabel, q1X, q1Y);

    // Quadrant 2 (top-left)
    ctx.fillText(topLeftLabel, q2X, q2Y);

    // Quadrant 3 (bottom-left)
    ctx.fillText(lowerLeftLabel, q3X, q3Y);

    // Quadrant 4 (bottom-right)
    ctx.fillText(lowerRightLabel, q4X, q4Y);

    ctx.restore();
  },
};

// Register Chart.js components
ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  Tooltip,
  Legend,
  Filler,
  watermarkPlugin
);

export interface IListItem {
  Id: number;
  Title?: string;
  Resolvability?: string;
  Opportunity?: string;
  [key: string]: any;
}

export default class SpfxIssueDetails extends React.Component<
  ISpfxIssueDetailsProps,
  { items: IListItem[]; loading: boolean; error: string | undefined }
> {
  constructor(props: ISpfxIssueDetailsProps) {
    super(props);
    this.state = {
      items: [],
      loading: true,
      error: undefined,
    };
  }

  public componentDidMount(): void {
    void this._fetchListItems();

    // Log the ID parameter for debugging (optional)
    if (this.props.idFromUrl) {
      console.log("ID parameter from URL:", this.props.idFromUrl);
    }
  }

  public componentDidUpdate(prevProps: ISpfxIssueDetailsProps): void {
    if (
      prevProps.marketAccessIssueList !== this.props.marketAccessIssueList ||
      prevProps.xAxisMeasure !== this.props.xAxisMeasure ||
      prevProps.yAxisMeasure !== this.props.yAxisMeasure
    ) {
      void this._fetchListItems();
    }
    // Force chart update if watermark setting changed
    if (prevProps.showWatermark !== this.props.showWatermark) {
      this.forceUpdate();
    }
  }

  private _fetchListItems = async (): Promise<void> => {
    const {
      marketAccessIssueList,
      context,
      idFromUrl,
      xAxisMeasure,
      yAxisMeasure,
    } = this.props;

    if (!marketAccessIssueList) {
      this.setState({
        items: [],
        loading: false,
        error: "List name not specified",
      });
      return;
    }

    if (!xAxisMeasure || !yAxisMeasure) {
      this.setState({
        items: [],
        loading: false,
        error: "Please select X-Axis and Y-Axis measures",
      });
      return;
    }

    this.setState({ loading: true, error: undefined });

    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      // Select dynamic fields instead of hardcoded ones
      let listUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
        marketAccessIssueList
      )}')/items?$select=Id,Title,${xAxisMeasure},${yAxisMeasure}`;

      // If ID is provided, filter by that specific item
      if (idFromUrl) {
        listUrl += `&$filter=Id eq ${idFromUrl}`;
      }

      const response = await context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(
          `Failed to fetch list items: ${response.status} ${response.statusText}`
        );
      }

      const data = await response.json();
      this.setState({
        items: data.value || [],
        loading: false,
        error: undefined,
      });
    } catch (error) {
      this.setState({ items: [], loading: false, error: error.message });
    }
  };

  private _getQuadrantColor = (
    x: number,
    y: number
  ): { bg: string; border: string } => {
    // Quadrant 1 (top-right): x >= 25, y >= 25 - Green (high resolvability, high opportunity)
    // Quadrant 2 (top-left): x < 25, y >= 25 - Yellow (low resolvability, high opportunity)
    // Quadrant 3 (bottom-left): x < 25, y < 25 - Red (low resolvability, low opportunity)
    // Quadrant 4 (bottom-right): x >= 25, y < 25 - Orange (high resolvability, low opportunity)

    if (x >= 25 && y >= 25) {
      return { bg: "rgba(75, 192, 192, 0.6)", border: "rgba(75, 192, 192, 1)" }; // Green - Q1
    } else if (x < 25 && y >= 25) {
      return { bg: "rgba(255, 206, 86, 0.6)", border: "rgba(255, 206, 86, 1)" }; // Yellow - Q2
    } else if (x < 25 && y < 25) {
      return { bg: "rgba(255, 99, 132, 0.6)", border: "rgba(255, 99, 132, 1)" }; // Red - Q3
    } else {
      return { bg: "rgba(255, 159, 64, 0.6)", border: "rgba(255, 159, 64, 1)" }; // Orange - Q4
    }
  };

  private _prepareChartData = () => {
    const { items } = this.state;
    const { xAxisMeasure, yAxisMeasure } = this.props;

    // Filter items that have valid X and Y axis values
    const chartData = items
      .filter((item) => {
        const xValue = this._parseNumber(item[xAxisMeasure]);
        const yValue = this._parseNumber(item[yAxisMeasure]);
        return xValue !== null && yValue !== null;
      })
      .map((item) => {
        const x = this._parseNumber(item[xAxisMeasure]) as number;
        const y = this._parseNumber(item[yAxisMeasure]) as number;
        const colors = this._getQuadrantColor(x, y);
        return {
          x: x,
          y: y,
          label: item.Title || `Item ${item.Id}`,
          backgroundColor: colors.bg,
          borderColor: colors.border,
        };
      });

    return chartData;
  };

  private _parseNumber = (value: any): number | null => {
    if (value === null || value === undefined || value === "") {
      return null;
    }
    const num = typeof value === "string" ? parseFloat(value) : Number(value);
    return isNaN(num) ? null : num;
  };

  private _getFieldDisplayName = (internalName: string): string => {
    if (!internalName) return "";
    // Convert internal name to display name (e.g., "Resolvability" -> "Resolvability")
    // If it contains underscores, convert to title case
    return internalName
      .split("_")
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
      .join(" ");
  };

  private _getChartOptions = () => {
    const {
      xAxisMeasure,
      yAxisMeasure,
      showWatermark,
      topRightLabel,
      topLeftLabel,
      lowerRightLabel,
      lowerLeftLabel,
    } = this.props;

    // Get field titles for display (fallback to internal name)
    const xAxisLabel = this._getFieldDisplayName(xAxisMeasure);
    const yAxisLabel = this._getFieldDisplayName(yAxisMeasure);
    const chartTitle = `${xAxisLabel} vs ${yAxisLabel}`;

    return {
      responsive: true,
      maintainAspectRatio: true,
      aspectRatio: 1,
      resizeDelay: 0,
      scales: {
        x: {
          type: "linear" as const,
          position: "bottom" as const,
          min: 0,
          max: 50,
          title: {
            display: true,
            text: xAxisLabel,
          },
          ticks: {
            stepSize: 25,
            callback: function (value: any) {
              // Only show labels at 0, 25, 50
              if (value === 0 || value === 25 || value === 50) {
                return value;
              }
              return "";
            },
          },
          grid: {
            color: (context: any) => {
              if (context.tick.value === 25) {
                return "rgba(0, 0, 0, 0.3)"; // Darker line for quadrant divider
              }
              return "rgba(0, 0, 0, 0.1)";
            },
            lineWidth: (context: any) => {
              if (context.tick.value === 25) {
                return 2; // Thicker line for quadrant divider
              }
              return 1;
            },
          },
        },
        y: {
          type: "linear" as const,
          min: 0,
          max: 50,
          title: {
            display: true,
            text: yAxisLabel,
          },
          ticks: {
            stepSize: 25,
            callback: function (value: any) {
              // Only show labels at 0, 25, 50
              if (value === 0 || value === 25 || value === 50) {
                return value;
              }
              return "";
            },
          },
          grid: {
            color: (context: any) => {
              if (context.tick.value === 25) {
                return "rgba(0, 0, 0, 0.3)"; // Darker line for quadrant divider
              }
              return "rgba(0, 0, 0, 0.1)";
            },
            lineWidth: (context: any) => {
              if (context.tick.value === 25) {
                return 2; // Thicker line for quadrant divider
              }
              return 1;
            },
          },
        },
      },
      plugins: {
        title: {
          display: true,
          text: chartTitle,
        },
        legend: {
          display: false,
        },
        tooltip: {
          callbacks: {
            label: (context: any) => {
              const point = context.raw;
              return `${point.label || "Item"}: (${point.x}, ${point.y})`;
            },
          },
        },
        watermark: {
          enabled: showWatermark,
          labels: {
            topRight: topRightLabel,
            topLeft: topLeftLabel,
            lowerRight: lowerRightLabel,
            lowerLeft: lowerLeftLabel,
          },
        },
      },
    };
  };

  private _getChartSize = (): number => {
    // Convert slider value (1-10) to chart width in pixels
    // 1 = 380px, 10 = 1000px (largest)
    const { chartSize } = this.props;
    const minSize = 380;
    const maxSize = 1000;
    // Map 1-10 to 380-1000: subtract 1 to make it 0-9, then divide by 9 to normalize
    return minSize + ((chartSize - 1) / 9) * (maxSize - minSize);
  };

  private _getInnerChartSize = (): number => {
    // Calculate inner chart size accounting for padding (20px on each side = 40px total)
    const containerSize = this._getChartSize();
    return containerSize - 40; // Subtract padding
  };

  public render(): React.ReactElement<ISpfxIssueDetailsProps> {
    const { loading, error } = this.state;

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
          <div style={{ color: "red" }}>Error: {error}</div>
        </section>
      );
    }

    const chartData = this._prepareChartData();
    const containerWidth = this._getChartSize();
    const chartWidth = this._getInnerChartSize();

    return (
      <section className={styles.spfxIssueDetails}>
        {chartData.length > 0 ? (
          <div
            style={{
              marginTop: "20px",
              padding: "20px",
              border: "1px solid #ddd",
              backgroundColor: "#fff",
              width: `${containerWidth}px`,
              marginLeft: "auto",
              marginRight: "auto",
              boxSizing: "border-box",
            }}
          >
            <div
              style={{
                width: `${chartWidth}px`,
                height: `${chartWidth}px`,
                margin: "0 auto",
              }}
            >
              <Scatter
                key={`chart-${this.props.chartSize}`}
                data={{
                  datasets: [
                    {
                      label: "Items",
                      data: chartData.map((point) => ({
                        x: point.x,
                        y: point.y,
                        label: point.label,
                      })),
                      pointBackgroundColor: chartData.map(
                        (point) => point.backgroundColor
                      ),
                      pointBorderColor: chartData.map(
                        (point) => point.borderColor
                      ),
                      pointRadius: 6,
                      pointHoverRadius: 8,
                    },
                  ],
                }}
                options={this._getChartOptions()}
              />
            </div>
          </div>
        ) : (
          <div style={{ padding: "20px", textAlign: "center" }}>
            No data available for chart. Please ensure X-Axis and Y-Axis
            measures are selected and the list contains data.
          </div>
        )}
      </section>
    );
  }
}
