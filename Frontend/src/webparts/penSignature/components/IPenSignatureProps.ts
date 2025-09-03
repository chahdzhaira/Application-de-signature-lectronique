import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPenSignatureProps {
   context : WebPartContext
   description?: string;
   isDarkTheme?: boolean;
   environmentMessage?: string;
   hasTeamsContext?: boolean;
   userDisplayName?: string;

   onClose?: () => void;
}
