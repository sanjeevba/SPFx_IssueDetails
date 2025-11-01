import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpfxIssueDetailsProps {
  description: string;
  marketAccessIssueList: string;
  xAxisMeasure: string;
  yAxisMeasure: string;
  chartSize: number;
  showWatermark: boolean;
  topRightLabel: string;
  topLeftLabel: string;
  lowerRightLabel: string;
  lowerLeftLabel: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  idFromUrl?: string | null;
}
