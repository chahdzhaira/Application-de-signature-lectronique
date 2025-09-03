import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHomeProps {
  context: WebPartContext
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}
