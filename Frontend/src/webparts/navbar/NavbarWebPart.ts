import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import Navbar from './components/Navbar';
import { sp } from "@pnp/sp/presets/all"



export default class NavbarWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    const element: React.ReactElement = React.createElement(
      Navbar,
      {
        context: this.context
      }
    );
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
