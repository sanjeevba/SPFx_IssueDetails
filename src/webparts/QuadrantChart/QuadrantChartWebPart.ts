import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import * as strings from "QuadrantChartWebPartStrings";
import QuadrantChart from "./components/QuadrantChart";
import { IQuadrantChartProps } from "./components/IQuadrantChartProps";

export interface IQuadrantChartWebPartProps {
  description: string;
  marketAccessIssueList: string;
  xAxisMeasure: string;
  yAxisMeasure: string;
  chartSize: number;
  showWatermark: boolean;
  topRightLabel: string;
  topLeftLabel: string;
  lowerRightLabel: string;
  lowerLeftLabel: string;
}

export default class QuadrantChartWebPart extends BaseClientSideWebPart<IQuadrantChartWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private _availableLists: { key: string; text: string }[] = [];
  private _numericColumns: { key: string; text: string }[] = [];

  private _getUrlParameter(paramName: string): string | undefined {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(paramName) || undefined;
  }

  public render(): void {
    // Read the ID parameter from URL
    const idFromUrl = this._getUrlParameter("ID");

    const element: React.ReactElement<IQuadrantChartProps> =
      React.createElement(QuadrantChart, {
        description: this.properties.description,
        marketAccessIssueList: this.properties.marketAccessIssueList,
        xAxisMeasure: this.properties.xAxisMeasure || "",
        yAxisMeasure: this.properties.yAxisMeasure || "",
        chartSize: this.properties.chartSize ?? 5,
        showWatermark: this.properties.showWatermark ?? true,
        topRightLabel: this.properties.topRightLabel || "1 - High Priority",
        topLeftLabel: this.properties.topLeftLabel || "2O - Big Impact",
        lowerRightLabel: this.properties.lowerRightLabel || "2R - Quick Win",
        lowerLeftLabel: this.properties.lowerLeftLabel || "3 - Low Priority",
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

  protected async loadPropertyPaneResources(): Promise<void> {
    await this._fetchAvailableLists();
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._fetchAvailableLists()
      .then(() => {
        this.context.propertyPane.refresh();
      })
      .catch(() => {
        // Error already logged in _fetchAvailableLists
      });

    // Fetch numeric columns if a list is selected
    if (this.properties.marketAccessIssueList) {
      this._fetchNumericColumns(this.properties.marketAccessIssueList)
        .then(() => {
          this.context.propertyPane.refresh();
        })
        .catch(() => {
          // Error already logged in _fetchNumericColumns
        });
    }
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    // If the list selection changed, fetch the numeric columns for that list
    if (propertyPath === "marketAccessIssueList" && newValue) {
      this._fetchNumericColumns(newValue)
        .then(() => {
          this.context.propertyPane.refresh();
        })
        .catch(() => {
          // Error already logged in _fetchNumericColumns
        });
    }

    // If the watermark toggle changed, refresh to show/hide label fields
    if (propertyPath === "showWatermark") {
      this.context.propertyPane.refresh();
    }
  }

  private _fetchAvailableLists = async (): Promise<void> => {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const listsUrl = `${webUrl}/_api/web/lists?$filter=Hidden eq false and BaseTemplate eq 100&$select=Id,Title&$orderby=Title`;

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          listsUrl,
          SPHttpClient.configurations.v1
        );

      if (response.ok) {
        const data = await response.json();
        this._availableLists = (data.value || []).map(
          (list: { Id: string; Title: string }) => ({
            key: list.Title,
            text: list.Title,
          })
        );
      } else {
        this._availableLists = [];
      }
    } catch (error) {
      console.error("Error fetching lists:", error);
      this._availableLists = [];
    }
  };

  private _fetchNumericColumns = async (listTitle: string): Promise<void> => {
    if (!listTitle) {
      this._numericColumns = [];
      return;
    }

    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const listUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
        listTitle
      )}')/fields?$filter=ReadOnlyField eq false and Hidden eq false and (TypeAsString eq 'Number' or TypeAsString eq 'Currency' or TypeAsString eq 'Decimal')&$select=InternalName,Title&$orderby=Title`;

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          listUrl,
          SPHttpClient.configurations.v1
        );

      if (response.ok) {
        const data = await response.json();
        this._numericColumns = (data.value || []).map(
          (field: { InternalName: string; Title: string }) => ({
            key: field.InternalName,
            text: field.Title || field.InternalName,
          })
        );
      } else {
        this._numericColumns = [];
      }
    } catch (error) {
      console.error("Error fetching numeric columns:", error);
      this._numericColumns = [];
    }
  };

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
                PropertyPaneDropdown("marketAccessIssueList", {
                  label: "List name",
                  options:
                    this._availableLists.length > 0
                      ? this._availableLists
                      : [{ key: "", text: "Loading lists..." }],
                  selectedKey: this.properties.marketAccessIssueList || "",
                }),
                PropertyPaneDropdown("xAxisMeasure", {
                  label: "X-Axis Measure",
                  options:
                    this._numericColumns.length > 0
                      ? this._numericColumns
                      : [
                          {
                            key: "",
                            text: this.properties.marketAccessIssueList
                              ? "Loading columns..."
                              : "Select a list first",
                          },
                        ],
                  selectedKey: this.properties.xAxisMeasure || "",
                  disabled: !this.properties.marketAccessIssueList,
                }),
                PropertyPaneDropdown("yAxisMeasure", {
                  label: "Y-Axis Measure",
                  options:
                    this._numericColumns.length > 0
                      ? this._numericColumns
                      : [
                          {
                            key: "",
                            text: this.properties.marketAccessIssueList
                              ? "Loading columns..."
                              : "Select a list first",
                          },
                        ],
                  selectedKey: this.properties.yAxisMeasure || "",
                  disabled: !this.properties.marketAccessIssueList,
                }),
                PropertyPaneSlider("chartSize", {
                  label: "Chart Size",
                  min: 1,
                  max: 10,
                  value: this.properties.chartSize ?? 5,
                  showValue: true,
                }),
                PropertyPaneToggle("showWatermark", {
                  label: "Show Watermark",
                  checked: this.properties.showWatermark ?? true,
                }),
                ...(this.properties.showWatermark ?? true
                  ? [
                      PropertyPaneTextField("topRightLabel", {
                        label: "Top Right Label",
                        value:
                          this.properties.topRightLabel || "1 - High Priority",
                      }),
                      PropertyPaneTextField("topLeftLabel", {
                        label: "Top Left Label",
                        value:
                          this.properties.topLeftLabel || "2O - Big Impact",
                      }),
                      PropertyPaneTextField("lowerRightLabel", {
                        label: "Lower Right Label",
                        value:
                          this.properties.lowerRightLabel || "2R - Quick Win",
                      }),
                      PropertyPaneTextField("lowerLeftLabel", {
                        label: "Lower Left Label",
                        value:
                          this.properties.lowerLeftLabel || "3 - Low Priority",
                      }),
                    ]
                  : []),
              ],
            },
          ],
        },
      ],
    };
  }
}
