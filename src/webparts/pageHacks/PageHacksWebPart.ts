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
  hideMargin: boolean;
  hideNav: boolean;
  hidePadding: boolean;
  hideSocial: boolean;
  isFullWidth: boolean;
}

export default class PageHacksWebPart extends BaseClientSideWebPart<IPageHacksWebPartProps> {

  public render(): void {
    // Clear this element
    while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

    // Update the canvas padding
    this.properties.hideMargin ? this.generateStyle(`div[data-automation-id='CanvasControl'] { margin: 0 !important; }`) : null;

    // Update the canvas padding
    this.properties.hidePadding ? this.generateStyle(`div[data-automation-id='CanvasControl'] { padding: 0 !important; }`) : null;

    // Update the canvas width
    this.properties.isFullWidth ? this.generateStyle(`div[data-automation-id='CanvasZone'] > div:first-child { max-width: none !important; }`) : null;

    // Update the page header
    this.properties.hideHeader ? this.generateStyle(`div[data-automation-id='pageHeader'] { display: none !important; }`) : null;

    // Update the page navigation
    this.properties.hideNav ? this.generateStyle(`#spLeftNav { display: none !important; }`) : null;

    // Update the page social footer
    this.properties.hideSocial ? this.generateStyle(`#CommentsWrapper { display: none !important; }`) : null;

    // Render the button if we are in edit mode
    if (this.displayMode === DisplayMode.Edit) {
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
              groupName: strings.SettingsGroupName,
              groupFields: [
                PropertyPaneToggle('hideMargin', {
                  label: strings.CanvasMarginFieldLabel,
                  offText: strings.CanvasMarginFieldOffText,
                  onText: strings.CanvasMarginFieldOnText
                }),
                PropertyPaneToggle('hidePadding', {
                  label: strings.CanvasPaddingFieldLabel,
                  offText: strings.CanvasPaddingFieldOffText,
                  onText: strings.CanvasPaddingFieldOnText
                }),
                PropertyPaneToggle('isFullWidth', {
                  label: strings.CanvasWidthFieldLabel,
                  offText: strings.CanvasWidthFieldOffText,
                  onText: strings.CanvasWidthFieldOnText
                }),
                PropertyPaneToggle('hideHeader', {
                  label: strings.PageHeaderFieldLabel,
                  offText: strings.PageHeaderFieldOffText,
                  onText: strings.PageHeaderFieldOnText
                }),
                PropertyPaneToggle('hideNav', {
                  label: strings.PageNavigationFieldLabel,
                  offText: strings.PageNavigationFieldOffText,
                  onText: strings.PageNavigationFieldOnText
                }),
                PropertyPaneToggle('hideSocial', {
                  label: strings.PageSocialFieldLabel,
                  offText: strings.PageSocialFieldOffText,
                  onText: strings.PageSocialFieldOnText
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
              groupName: strings.AboutGroupName,
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

  // Generate a style tag
  private generateStyle(innerHTML: string): void {
    const style = document.createElement("style");
    style.innerHTML = innerHTML;
    this.domElement.appendChild(style);
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
}