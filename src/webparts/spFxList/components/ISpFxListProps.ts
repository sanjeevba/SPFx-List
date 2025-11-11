import { SPHttpClient } from '@microsoft/sp-http';

export interface ISpFxListProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  selectedListId: string;
  spHttpClient: SPHttpClient;
  webUrl: string;
}
