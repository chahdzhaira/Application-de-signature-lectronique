import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ChooseSignature from './components/ChooseSignature';

export default class ChooseSignatureWebPart extends BaseClientSideWebPart<{}> {


  public render(): void {
    const element: React.ReactElement = React.createElement(
      ChooseSignature,
      {
        context: this.context
      }
    );
    
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

}
