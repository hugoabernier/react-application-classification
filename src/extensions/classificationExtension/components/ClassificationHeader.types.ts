import { ExtensionContext } from "@microsoft/sp-extension-base";

export interface IClassificationHeaderProps {
    context: ExtensionContext;
    ClassificationPropertyBag: string;
    DefaultClassification: string;
    DefaultHandlingUrl: string;
}

export interface IClassificationHeaderState {
    isLoading: boolean;
    businessImpact: string;
}

// change this value to whatever you want to use as the property bag value name
// export const ClassificationPropertyBag: string = "sc_x005f_BusinessImpact";

// change this value to whatever you want the default classification you wish to use. "undefined" means no default.
// export const DefaultClassification: string = undefined;

// export const DefaultHandlingUrl: string = "/SitePages/Handling-instructions.aspx";