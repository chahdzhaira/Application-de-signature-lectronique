import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import Footer from "./components/Footer";

export interface IFooterWebPartProps {
  description: string;
}

export default class FooterWebPart extends BaseClientSideWebPart<IFooterWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IFooterWebPartProps> = React.createElement(Footer);

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }


}
