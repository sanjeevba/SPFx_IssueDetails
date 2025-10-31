import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISpfxIssueDetailsProps {
  description: string;
  marketAccessIssueList: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
