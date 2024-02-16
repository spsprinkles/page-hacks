import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneButton, PropertyPaneHorizontalRule, PropertyPaneLink, PropertyPaneLabel, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as common from './common';
import * as strings from 'PageHacksWebPartStrings';
import './PageHacksWebPart.module.scss';

import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";

export interface IPageHacksWebPartProps {
  hideHeader: boolean;
  hideNav: boolean;
  hidePadding: boolean;
  isFullWidth: boolean;
}

export default class PageHacksWebPart extends BaseClientSideWebPart<IPageHacksWebPartProps> {

  public render(): void {
    // Update the page header
    this.updatePageHeader(this.properties.hideHeader);

    // Update the page navigation
    this.updatePageNavigation(this.properties.hideNav);

    // Update the page padding
    this.updatePadding(this.properties.hidePadding);

    // Update the page width
    this.updatePageWidth(this.properties.isFullWidth);

    // Do nothing if we are in display mode
    if (this.displayMode === DisplayMode.Read) { return; }

    // Clear this element
    while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

    // Display a button to configure the page
    Components.Tooltip({
      el: this.domElement,
      content: "Configure Settings",
      btnProps: {
        className: "p-1 pe-2",
        iconType: common.getLogo(24, 24, "me-2"),
        text: "Settings",
        type: Components.ButtonTypes.OutlinePrimary,
        onClick: () => {
          // Show the edit panel
          this.context.propertyPane.open();
        }
      }
    });
  }

  protected get dataVersion(): Version {
    return Version.parse(this.context.manifest.version);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Settings:",
              groupFields: [
                PropertyPaneToggle('hidePadding', {
                  label: strings.PagePaddingFieldLabel,
                  offText: "Standard Padding",
                  onText: "Hide Padding"
                }),
                PropertyPaneToggle('isFullWidth', {
                  label: strings.PageWidthFieldLabel,
                  offText: "Standard Width",
                  onText: "Full Page Width"
                }),
                PropertyPaneToggle('hideHeader', {
                  label: strings.PageHeaderFieldLabel,
                  offText: "Show Page Header",
                  onText: "Hide Page Header"
                }),
                PropertyPaneToggle('hideNav', {
                  label: strings.PageNavigationFieldLabel,
                  offText: "Show Page Navigation",
                  onText: "Hide Page Navigation"
                }),
                PropertyPaneButton('', {
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
        },
        {
          groups: [
            {
              groupName: "About this app:",
              groupFields: [
                PropertyPaneLabel('version', {
                  text: "Version: " + this.context.manifest.version
                }),
                PropertyPaneLabel('description', {
                  text: strings.PropertyPaneDescription
                }),
                PropertyPaneLabel('about', {
                  text: "We think adding sprinkles to a donut just makes it better! SharePoint Sprinkles builds apps that are sprinkled on top of SharePoint, making your experience even better. Check out our site below to discover other SharePoint Sprinkles apps, or connect with us on GitHub."
                }),
                PropertyPaneLabel('support', {
                  text: "Are you having a problem or do you have a great idea for this app? Visit our GitHub link below to open an issue and let us know!"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLink('supportLink', {
                  href: "https://www.spsprinkles.com/",
                  text: "SharePoint Sprinkles",
                  target: "_blank"
                }),
                PropertyPaneLink('sourceLink', {
                  href: "https://github.com/spsprinkles/page-hacks/",
                  text: "View Source on GitHub",
                  target: "_blank"
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onInit(): Promise<void> {
    // Set the context
    ContextInfo.setPageContext(this.context.pageContext);

    // Return a promise
    return Promise.resolve();
  }

  protected onPropertyPaneRendered(): void {
    const setLogo = setInterval(() => {
      let closeBtn = document.querySelectorAll("div.spPropertyPaneContainer div[aria-label='Page Hacks property pane'] button[data-automation-id='propertyPaneClose']");
      if (closeBtn) {
        closeBtn.forEach((el: HTMLElement) => {
          let parent = el.parentElement;
          if (parent && !(parent.firstChild as HTMLElement).classList.contains("logo")) { parent.prepend(common.getLogo(28, 28, 'logo me-2')) }
        });
        clearInterval(setLogo);
      }
    }, 50);
  }

  // Updates the extra padding on the page
  private updatePadding(hidePadding: boolean): void {
    const setPadding = setInterval(() => {
      // Get the canvas control elements
      let elCanvasControl = document.querySelectorAll("div[data-automation-id='CanvasControl']");
      if (elCanvasControl) {
        elCanvasControl.forEach((el: HTMLElement) => {
          // Set or clear margin
          hidePadding ? el.style.margin = "0" : el.style.margin = "";
          // Set or clear padding
          hidePadding ? el.style.padding = "0" : el.style.padding = "";
        });
        clearInterval(setPadding);
      }
    }, 50);
  }

  // Updates the page header visibility
  private updatePageHeader(hideHeader: boolean): void {
    const setPageHeader = setInterval(() => {
      // Get the header element
      let elHeader: HTMLElement = document.querySelector("div[data-automation-id='pageHeader']");
      if (elHeader) {
        // Update the header visibility
        elHeader.style.display = hideHeader ? "none" : "";
        clearInterval(setPageHeader);
      }
    }, 50);
  }

  // Updates the page navigation visibility
  private updatePageNavigation(hideNav: boolean): void {
    const setPageNav = setInterval(() => {
      // Get the navigation element
      let elNav: HTMLElement = document.querySelector("#spLeftNav");
      if (elNav) {
        // Update the navigation visibility
        elNav.style.display = hideNav ? "none" : "";
        clearInterval(setPageNav);
      }
    }, 50);
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
          let pageLayout = (item as any)["PageLayoutType"];

          // Set the header
          Modal.setHeader("Update Page Layout");

          // See if an item exists
          if (item) {
            // Display a form
            let form = Components.Form({
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
                    let selectedValue: Components.ICheckboxGroupItem = form.getValues()["PageLayout"];

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
                        let ctrl = form.getControl("PageLayout");
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
    const setPageWidth = setInterval(() => {
      // Get the canvas zone elements
      let elCanvasZone = document.querySelectorAll("div[data-automation-id='CanvasZone']");
      if (elCanvasZone) {
        elCanvasZone.forEach((el: HTMLElement) => {
          // Get the first child element
          let elCanvasChild = el.firstChild as HTMLElement;
          // Set or clear max width
          setFullWidth ? elCanvasChild.style.maxWidth = "none" : elCanvasChild.style.maxWidth = "";
        });
        clearInterval(setPageWidth);
      }
    }, 50);
  }
}