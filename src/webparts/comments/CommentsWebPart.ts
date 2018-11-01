import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CommentsWebPartStrings';
import Comments from './components/Comments';
import { ICommentsProps } from './components/ICommentsProps';

export interface ICommentsWebPartProps {
  description: string;
}

export default class CommentsWebPart extends BaseClientSideWebPart<ICommentsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICommentsProps > = React.createElement(
      Comments,
      {
        description: this.properties.description
        // test
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
