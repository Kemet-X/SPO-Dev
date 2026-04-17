import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import GraphLists from './components/GraphLists';
import type { IGraphListsProps } from './components/IGraphListsProps';

export default class GraphListsWebPart extends BaseClientSideWebPart<Record<string, never>> {
  public render(): void {
    const element: React.ReactElement<IGraphListsProps> = React.createElement(GraphLists, {
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
