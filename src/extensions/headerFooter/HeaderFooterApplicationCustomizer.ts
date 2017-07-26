import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
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
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HeaderFooterApplicationCustomizer
  extends BaseApplicationCustomizer<IHeaderFooterApplicationCustomizerProperties> {

  private customHeader: any;
  private customFooter: any;

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
          this.customHeader = $("<div id='customHeader' class='ms-dialogHidden' style='background-color:" + r.AllProperties.CustomSiteHeaderBgColor + ";color:" + r.AllProperties.CustomSiteHeaderColor + ";padding:3px;text-align:center;font-family:Segoe UI'><b>" + r.AllProperties.CustomSiteHeaderText + "</b></div>");
        }
        if (r.AllProperties.CustomSiteFooterEnabled == "true") {
          this.customFooter = $("<div id='customFooter' class='ms-dialogHidden' style='background-color:" + r.AllProperties.CustomSiteFooterBgColor + ";color:" + r.AllProperties.CustomSiteFooterColor + ";padding:3px;text-align:center;font-family:Segoe UI'><b>" + r.AllProperties.CustomSiteFooterText + "</b></div>");
        }

        resolve();
      });
    });
  }

  @override
  public onRender(): void {
    if ($("#spoAppComponent").length == 1) {
      // Site contents, List/library view
      this.customHeader.insertBefore("#spoAppComponent");
      this.customFooter.insertAfter("#spoAppComponent");
    } else {
      // Site page
      this.customHeader.insertBefore(".SPPageChrome");
      this.customFooter.insertAfter(".SPPageChrome");
    }

    $(window).resize(this.calcFooter);

    this.calcFooter(); 
  }

  private calcFooter(): void {
    var $footer = $("#customFooter");
    var footerheight = $footer.outerHeight();

    var $header = $("#customHeader");	
    var headerheight = $header.outerHeight();

    var $bodySelector;
    if ($("#spoAppComponent").length == 1) {
        // Site contents, List/library view
        $bodySelector = $("#spoAppComponent");
    } else {
        // Site page
        $bodySelector = $(".SPPageChrome");
    }
    
    var windowheight = $(window).height();
          
    // Resize 
    $bodySelector.css('height', windowheight - footerheight - headerheight);
  }
}