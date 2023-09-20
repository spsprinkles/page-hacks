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
  hideNav: boolean;
  isFullWidth: boolean;
}

export default class PageHacksWebPart extends BaseClientSideWebPart<IPageHacksWebPartProps> {

  public render(): void {
    // Update the page navigation
    this.updatePageNavigation(this.properties.hideNav);

    // Update the page width
    this.updatePageWidth(this.properties.isFullWidth);

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
              groupFields: [
                PropertyPaneToggle("hideNav", {
                  label: strings.PageNavigationFieldLabel,
                  offText: strings.PageNavigationFieldDescription,
                  onText: strings.PageNavigationFieldDescription
                }),
                PropertyPaneToggle("isFullWidth", {
                  label: strings.PageWidthFieldLabel,
                  offText: strings.PageWidthFieldDescription,
                  onText: strings.PageWidthFieldDescription
                }),
                PropertyPaneButton("", {
                  description: strings.PageTypeFieldDescription,
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

  // Updates the page navigation visibility
  private updatePageNavigation(hideNav: boolean) {
    // Get the navigation element
    const elNav: HTMLElement = document.querySelector("#spLeftNav");
    if (elNav) {
      // Update the navigation visibility
      elNav.style.display = hideNav ? "none" : "";
    }
  }

  // Displays the panel to update the page template
  private updatePageTemplate(): void {
    // Display a loading dialog
    LoadingDialog.setHeader("Getting Page Information");
    LoadingDialog.setBody("This will close after the page information is loaded...");
    LoadingDialog.show();

    // Clear the modal
    Modal.clear();

    // Ensure the page context information exists
    if (this.context.pageContext.list && this.context.pageContext.listItem) {
      // Get the page information
      Web().Lists(this.context.pageContext.list.title).Items(this.context.pageContext.listItem.id).query({
        Select: ["PageLayoutType"]
      }).execute(
        // Success
        item => {
          const pageLayout = (item as any)["PageLayoutType"];

          // Set the header
          Modal.setHeader("Update Page Layout");

          // See if an item exists
          if (item) {
            // Display a form
            const form = Components.Form({
              el: Modal.BodyElement,
              controls: [{
                label: "Select a Layout",
                name: "PageLayout",
                description: "Select a page layout",
                type: Components.FormControlTypes.Switch,
                value: pageLayout,
                required: true,
                items: [
                  {
                    label: "Article"
                  },
                  {
                    label: "Home"
                  },
                  {
                    label: "SingleWebPartAppPage"
                  }
                ],
                onValidate: (ctrl, results) => {
                  // Ensure a value exists
                  if (results.value) {
                    // See if this layout is already set
                    if ((results.value as Components.ICheckboxGroupItem).label === pageLayout) {
                      // Set the flag
                      results.isValid = false;
                      results.invalidMessage = "The page is already set to this layout, please select a different layout."
                    }
                  } else {
                    // Set the flag
                    results.isValid = false;
                    results.invalidMessage = "Please select a page layout type."
                  }

                  // Return the results
                  return results;
                }
              } as Components.IFormControlPropsSwitch]
            });

            // Set the footer
            Components.Tooltip({
              el: Modal.FooterElement,
              content: "Click to update the page template.",
              btnProps: {
                text: "Update",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                  // Validate the form
                  if (form.isValid()) {
                    const selectedValue: Components.ICheckboxGroupItem = form.getValues()["PageLayout"];

                    // Update the item
                    item.update({
                      PageLayoutType: selectedValue.label
                    }).execute(
                      () => {
                        // Refresh the page
                        location.reload();
                      },
                      () => {
                        // Set the validation
                        const ctrl = form.getControl("PageLayout");
                        ctrl.updateValidation(ctrl.el, {
                          invalidMessage: "Error updating the item. Please refresh and try again.",
                          isValid: false
                        });
                      }
                    );
                  }
                }
              }
            })
          } else {
            Modal.setBody("Unable to get the item for this page.");
          }

          // Display a modal
          Modal.show();

          // Hide the dialog
          LoadingDialog.hide();
        },
        // Error
        () => {
          // Unable to determine the page information
          LoadingDialog.hide();

          // Display an error message
          Modal.setHeader("Error");
          Modal.setBody("There was an error getting the page information...");
          Modal.show();
        });
    } else {
      // Unable to determine the page information
      LoadingDialog.hide();

      // Display an error message
      Modal.setHeader("Error");
      Modal.setBody("Unable to get the page information from the context. This is not a modern site page...");
      Modal.show();
    }
  }

  // Updates the page width
  private updatePageWidth(setFullWidth: boolean): void {
    // Get the canvas zone element
    const elCanvas = document.querySelector(".CanvasZone.row");
    if (elCanvas) {
      // See if we are setting the full width
      if (setFullWidth) {
        elCanvas.classList.add("CanvasZone--fullWidth");
        elCanvas.classList.add("CanvasZone--fullWidth--read");
      } else {
        elCanvas.classList.remove("CanvasZone--fullWidth");
        elCanvas.classList.remove("CanvasZone--fullWidth--read");
      }
    }
  }
}