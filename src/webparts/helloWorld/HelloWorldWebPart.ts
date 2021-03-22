import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/comments";
import "@pnp/sp/hubsites";
import { IHubSiteInfo } from "@pnp/sp/hubsites";
import { IItem } from "@pnp/sp/items";


export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context,
        defaultCachingStore: "session",
        defaultCachingTimeoutSeconds: 300,
        enableCacheExpiration: true,
        globalCacheDisable: false,
      });
    });
  }

  public  async render(): Promise<void> {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description
      }
    );

    await this.getHubSiteFromService(this.context?.pageContext?.legacyPageContext?.hubSiteId);
    await this.getPageListItem((this.context?.pageContext?.list.id as any)._guid, this.context?.pageContext?.listItem.id);

    ReactDom.render(element, this.domElement);
  }

  private getHubSiteFromService = async (hubSiteId: string) => {
    try {
      console.log('Getting HUB SITE INFO');
      const hubSiteInfo: IHubSiteInfo = await sp.hubSites.getById(hubSiteId)();
      console.log('IHubSiteInfo from service - ', hubSiteInfo);
      return hubSiteInfo;
    }
    catch (error) {
      console.error('PnP call error - exception getting hub site - ', error);
      return undefined;
    }
  }

  private getPageListItem = async (listId: string, listItemId: number): Promise<IItem> => {
    let pageListItem: any;
      try {
        pageListItem = await sp.web.lists.getById(listId).items.getById(listItemId).fieldValuesAsHTML.get();

        let pageAuthorItem = await sp.web.lists.getById(listId).items.getById(listItemId).usingCaching()
          .select("Created", "Modified", "FirstPublishedDate", "Author/EMail", "Author/Title", "Editor/EMail", "Editor/Title")
          .expand("Editor", "Author").get();
      }
      catch (error) {
        console.warn('PnP call error list item - ', error);
      }

    
    console.info('PageListItem - ', pageListItem);

    return pageListItem;
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
