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
import "bootstrap/dist/css/bootstrap.min.css";

import { SPComponentLoader } from '@microsoft/sp-loader';

export interface ICommentsWebPartProps {
  queryStrItemIdFieldName: string;
  listName: string;
}

export default class CommentsWebPart extends BaseClientSideWebPart<ICommentsWebPartProps> {
 protected onInit(): Promise<void> {
   SPComponentLoader.loadCss(this.context.pageContext.web.absoluteUrl+'/Style%20Library/MarketPlace/CommentsStyle.css?csf=1&e=Tn9IJN');
   return super.onInit();
 }

  public render(): void {
    const element: React.ReactElement<ICommentsProps > = React.createElement(
      Comments,
      {
        queryStrItemIdFieldName: this.properties.queryStrItemIdFieldName,
        listName: this.properties.listName,
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
                PropertyPaneTextField('queryStrItemIdFieldName', {
                  label: strings.QueryStrItemIdFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
