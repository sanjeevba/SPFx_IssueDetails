import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpfxIssueDetailsProps {
  description: string;
  marketAccessIssueList: string;
  chartSize: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  idFromUrl?: string | null;
}
