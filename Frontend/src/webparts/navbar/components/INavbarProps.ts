import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INavbarProps {
  context : WebPartContext
  description?: string;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}
