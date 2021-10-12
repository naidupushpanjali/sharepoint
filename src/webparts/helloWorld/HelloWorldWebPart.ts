import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "HelloWorldWebPartStrings";
import HelloWorld from "./components/HelloWorld";
import { IHelloWorldProps } from "./components/IHelloWorldProps";
import MockHttpClient from "./MockHttpClient";
import {
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import {
  Environment,
  EnvironmentType,
} from "@microsoft/sp-core-library";

export interface IHelloWorldWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> =
      React.createElement(HelloWorld, {
        description: this.properties.description,
      });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(
      this.domElement
    );
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField(
                  "description",
                  {
                    label:
                      strings.DescriptionFieldLabel,
                  }
                ),
              ],
            },
          ],
        },
      ],
    };
  }
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get().then(
      (data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }
    ) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
    console.log("here   ", this.context);
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  private _renderListAsync(): void {
    // Local environment
    if (
      Environment.type === EnvironmentType.Local
    ) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    } else if (
      Environment.type ==
        EnvironmentType.SharePoint ||
      Environment.type ==
        EnvironmentType.ClassicSharePoint
    ) {
      this._getListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }
  private _renderList(items: ISPList[]): void {
    let html: string = "";
    items.forEach((item: ISPList) => {
      html += `
    <ul>
      <li>
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });

    const listContainer: Element =
      this.domElement.querySelector(
        "#spListContainer"
      );
    listContainer.innerHTML = html;
  }
}
