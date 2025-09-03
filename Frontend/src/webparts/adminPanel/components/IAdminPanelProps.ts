import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAdminPanelProps {
  context : WebPartContext
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}
