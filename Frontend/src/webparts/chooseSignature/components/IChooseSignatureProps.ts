import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IChooseSignatureProps {
    context : WebPartContext
    description?: string;
    isDarkTheme?: boolean;
    environmentMessage?: string;
    hasTeamsContext?: boolean;
    userDisplayName?: string;

    onClose?: () => void; 
}
