import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import SignatureRequestOverview from './components/SignatureRequestOverview';
import { ISignatureRequestOverviewProps } from './components/ISignatureRequestOverviewProps';
import { sp } from "@pnp/sp/presets/all"

export interface ISignatureRequestOverviewWebPartProps {
  description: string;
}

export default class SignatureRequestOverviewWebPart extends BaseClientSideWebPart<ISignatureRequestOverviewWebPartProps> {


  public render(): void {
    const element: React.ReactElement<ISignatureRequestOverviewProps> = React.createElement(
      SignatureRequestOverview,
      {
        context: this.context 
      });
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context as any
      })
    })
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

}
