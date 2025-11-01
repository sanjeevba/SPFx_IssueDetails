import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "SpfxIssueDetailsWebPartStrings";
import SpfxIssueDetails from "./components/SpfxIssueDetails";
import { ISpfxIssueDetailsProps } from "./components/ISpfxIssueDetailsProps";

export interface ISpfxIssueDetailsWebPartProps {
  description: string;
  marketAccessIssueList: string;
  chartSize: number;
}

export default class SpfxIssueDetailsWebPart extends BaseClientSideWebPart<ISpfxIssueDetailsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  private _getUrlParameter(paramName: string): string | null {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(paramName);
  }

  public render(): void {
    // Read the ID parameter from URL
    const idFromUrl = this._getUrlParameter("ID");

    const element: React.ReactElement<ISpfxIssueDetailsProps> =
      React.createElement(SpfxIssueDetails, {
        description: this.properties.description,
        marketAccessIssueList: this.properties.marketAccessIssueList,
        chartSize: this.properties.chartSize ?? 5,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        idFromUrl: idFromUrl,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: "Properties",
              groupFields: [
                PropertyPaneTextField("marketAccessIssueList", {
                  label: "Market Access Issue List",
                  value: "MA Issue Tmp",
                }),
                PropertyPaneSlider("chartSize", {
                  label: "Chart Size",
                  min: 1,
                  max: 10,
                  value: this.properties.chartSize ?? 5,
                  showValue: true,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
