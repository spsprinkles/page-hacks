import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PageHacksWebPartStrings';

import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";

export interface IPageHacksWebPartProps {
  isFullWidth: boolean;
}

export default class PageHacksWebPart extends BaseClientSideWebPart<IPageHacksWebPartProps> {

  public render(): void {
    // See if we are displaying full width
    if (this.properties.isFullWidth) {
      // TODO
    }

    // Do nothing if we are in display mode
    if (this.displayMode === DisplayMode.Read) { return; }

    // Clear this element
    while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

    // Display a button to configure the page
    Components.Tooltip({
      el: this.domElement,
      content: "Click to configure this page",
      btnProps: {
        text: "Configure",
        type: Components.ButtonTypes.OutlinePrimary,
        onClick: () => {
          // Show the edit panel
          this.context.propertyPane.open();
        }
      }
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('0.1');
  }

  protected onInit(): Promise<void> {
    // Set the context
    ContextInfo.setPageContext(this.context.pageContext);

    // Return a promise
    return Promise.resolve();
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
                PropertyPaneToggle("isFullWidth", {
                  label: strings.PageWidthFieldLabel
                }),
                PropertyPaneButton("", {
                  text: strings.PageTypeFieldLabel,
                  onClick: () => {
                    // Show a modal to set the page type
                    this.updatePageTemplate();
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // Displays the panel to update the page template
  private updatePageTemplate(): void {
    // Display a loading dialog
    LoadingDialog.setHeader("Getting Page Information");
    LoadingDialog.setBody("This will close after the page information is loaded...");
    LoadingDialog.show();

    // Get the page information
    Web().getFileByUrl(location.pathname).execute(page => {
      // Display a modal
      Modal.setHeader("Update Page Template");
      Modal.setBody("The page template is currently: " + (page as any)["Template"]);
      Modal.show();
    }, () => {
      // Unable to determine the page information
      LoadingDialog.hide();

      // Display an error message
      Modal.setHeader("Error");
      Modal.setBody("There was an error getting the page information...");
      Modal.show();
    })
  }
}
