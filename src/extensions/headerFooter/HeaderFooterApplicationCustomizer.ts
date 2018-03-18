import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'headerFooterStrings';

// Custom imports
import * as $ from 'jquery';
import pnp from 'sp-pnp-js';

const LOG_SOURCE: string = 'HeaderFooterApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHeaderFooterApplicationCustomizerProperties {
  Header: string;
  Footer: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HeaderFooterApplicationCustomizer
  extends BaseApplicationCustomizer<IHeaderFooterApplicationCustomizerProperties> {

  private _headerPlaceholder: PlaceholderContent | undefined;
  private _footerPlaceholder: PlaceholderContent | undefined;
  
  @override
  public onInit(): Promise<void> {
    pnp.setup({
      spfxContext: this.context
    });

    return new Promise<void>((resolve, reject) => {
      pnp.sp.web.select("AllProperties").expand("AllProperties").get().then(r => {
        console.log("CustomSiteHeaderEnabled: " + r.AllProperties.CustomSiteHeaderEnabled);
        console.log("CustomSiteHeaderBgColor: " + r.AllProperties.CustomSiteHeaderBgColor);
        console.log("CustomSiteHeaderColor: " + r.AllProperties.CustomSiteHeaderColor);
        console.log("CustomSiteHeaderText: " + r.AllProperties.CustomSiteHeaderText);
        console.log("CustomSiteFooterEnabled: " + r.AllProperties.CustomSiteFooterEnabled);
        console.log("CustomSiteFooterBgColor: " + r.AllProperties.CustomSiteFooterBgColor);
        console.log("CustomSiteFooterColor: " + r.AllProperties.CustomSiteFooterColor);
        console.log("CustomSiteFooterText: " + r.AllProperties.CustomSiteFooterText);     

        if (r.AllProperties.CustomSiteHeaderEnabled == "true") {
          this.properties.Header = "<div id='customHeader' class='ms-dialogHidden' style='background-color:" + r.AllProperties.CustomSiteHeaderBgColor + ";color:" + r.AllProperties.CustomSiteHeaderColor + ";padding:3px;text-align:center;font-family:Segoe UI'><b>" + r.AllProperties.CustomSiteHeaderText + "</b></div>";
        }
        if (r.AllProperties.CustomSiteFooterEnabled == "true") {
          this.properties.Footer = "<div id='customFooter' class='ms-dialogHidden' style='background-color:" + r.AllProperties.CustomSiteFooterBgColor + ";color:" + r.AllProperties.CustomSiteFooterColor + ";padding:3px;text-align:center;font-family:Segoe UI'><b>" + r.AllProperties.CustomSiteFooterText + "</b></div>";
        }

        // Added to handle possible changes on the existence of placeholders
        this.context.placeholderProvider.changedEvent.add(this, this.onRender);
        // Call render method for generating the needed HTML elements
        this.onRender();

        resolve();
      });
    });
  }

  @override
  public onRender(): void {
    console.log('CustomHeader.onRender()');
    console.log('Available placeholders: ',
      this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the header placeholder
    if (!this._headerPlaceholder) {
      this._headerPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        {
          onDispose: this._onDispose
        });

      // The extension should not assume that the expected placeholder is available.
      if (!this._headerPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this.properties && this.properties.Header && this._headerPlaceholder.domElement) {
        this._headerPlaceholder.domElement.innerHTML = this.properties.Header;
      }
    }

    // Handling the footer placeholder
    if (!this._footerPlaceholder) {
      this._footerPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        {
          onDispose: this._onDispose
        });

      // The extension should not assume that the expected placeholder is available.
      if (!this._footerPlaceholder) {
        console.error('The expected placeholder (Bottom) was not found.');
        return;
      }

      if (this.properties && this.properties.Footer && this._footerPlaceholder.domElement) {
        this._footerPlaceholder.domElement.innerHTML = this.properties.Footer;
      }
    }
  }

  private _onDispose(): void {
    console.log('[HeaderFooterApplicationCustomizer._onDispose] Disposed custom Top and Bottom placeholders.');
  }
}
