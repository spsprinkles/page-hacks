import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneHorizontalRule, PropertyPaneLink, PropertyPaneLabel, PropertyPaneToggle } from '@microsoft/sp-property-pane';
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
  private _pageLayoutTypeDisabled: boolean = true;
  private _pageLayoutType: string;

  public render(): void {
    // Clear this element
    while (this.domElement.firstChild) { this.domElement.removeChild(this.domElement.firstChild); }

    // Update the canvas padding
    this.properties.hideMargin ? this.generateStyle(`div[data-automation-id='CanvasControl'] { margin: 0 !important; }`) : null;

    // Update the canvas padding
    this.properties.hidePadding ? this.generateStyle(`div[data-automation-id='CanvasControl'] { padding: 0 !important; }`) : null;

    // Update the canvas width
    this.properties.isFullWidth ? this.generateStyle(`div[data-automation-id='CanvasZone'] > div[data-automation-id='CanvasZone-SectionContainer'] { max-width: none !important; }`) : null;

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
          text: "Page Hacks",
          type: Components.ButtonTypes.OutlinePrimary,
          onClick: () => {
            // Show the edit panel
            this.context.propertyPane.open();
          }
        }
      });

      // Get the current pageLayoutType
      this.getPageLayoutType().then((pageLayoutType) => {
        // Set the pageLayoutType value
        this._pageLayoutType = pageLayoutType;
        this._pageLayoutTypeDisabled = false;
      }, () => {
        // Disable the dropdown on error
        this._pageLayoutTypeDisabled = true;
      });
    } else {
      // Hide this webpart
      this.hideWebpart();
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
                PropertyPaneDropdown('', {
                  disabled: this._pageLayoutTypeDisabled,
                  label: strings.PageLayoutTypeFieldLabel,
                  selectedKey: this._pageLayoutType,
                  options: [
                    { key: "Article", text: "Article" },
                    { key: "Home", text: "Home" }
                  ]
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

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === "") {
      this.updatePageLayoutType(newValue);
    }
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

  // Get the current page layout type
  private getPageLayoutType(): PromiseLike<string> {
    // Return a promise
    return new Promise((resolve, reject) => {
      // Ensure the page context information exists
      if (this.context.pageContext.list && this.context.pageContext.listItem) {
        // Get the page information
        Web().Lists(this.context.pageContext.list.title).Items(this.context.pageContext.listItem.id).query({
          Select: ["PageLayoutType"]
        }).execute(
          // Success
          item => {
            // See if an item exists
            item ? resolve((item as any)["PageLayoutType"]) : reject;
          },
          // Error
          reject);
      } else {
        reject("Unable to get the page information from the context. This is not a modern site page...");
      }
    });
  }

  // Hides this webpart
  private hideWebpart(): void {
    // Find the parent div element and make it hidden
    let el = this.domElement.parentElement;
    while (el && el.getAttribute("data-automation-id") !== "CanvasControl") { el = el.parentElement; }

    // See if the element was found
    if (el) {
      // Hide the element
      el.style.display = "none";
    }
  }

  // Updates the page template
  private updatePageLayoutType(newValue: any): void {
    // Display a loading dialog
    LoadingDialog.setHeader("Getting Page Information");
    LoadingDialog.setBody("This will close after the page information is loaded...");
    LoadingDialog.show();

    // Clear the modal
    Modal.clear();

    // Hide the modal by default
    let showModal = false;

    // Ensure the page context information exists
    if (this.context.pageContext.list && this.context.pageContext.listItem) {
      // Get the page information
      Web().Lists(this.context.pageContext.list.title).Items(this.context.pageContext.listItem.id).query({
        Select: ["PageLayoutType"]
      }).execute(
        // Success
        item => {
          let pageLayoutType = (item as any)["PageLayoutType"];

          // Set the header
          Modal.setHeader("Update Page Layout");

          // See if an item exists
          if (item) {
            // If the current page layoutType matches the old value, perform the update using the newValue
            if (pageLayoutType === this._pageLayoutType) {
              // Update the item
              item.update({
                PageLayoutType: newValue
              }).execute(
                () => {
                  // Reload the page
                  window.location.reload();
                },
                () => {
                  // Error
                  Modal.setBody("Unable to save the new PageLayoutType to this page.");
                  showModal = true;
                }
              );
            } else {
              Modal.setBody("The PageLayoutType does not need to be updated on this page.");
              showModal = true;
            }
          } else {
            Modal.setBody("Unable to get the information needed from this page.");
            showModal = true;
          }

          // Hide the dialog
          LoadingDialog.hide();

          // Display a modal
          showModal ? Modal.show() : null;
        },
        // Error
        () => {
          // Unable to determine the page information
          LoadingDialog.hide();

          // Display an error message
          Modal.setHeader("Error");
          Modal.setBody("There was an error getting the page information...");
          showModal = true;
          showModal ? Modal.show() : null;
        });
    } else {
      // Unable to determine the page information
      LoadingDialog.hide();

      // Display an error message
      Modal.setHeader("Error");
      Modal.setBody("Unable to get the page information from the context. This is not a modern site page...");
      showModal = true;
      showModal ? Modal.show() : null;
    }
  }
}