import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import {
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'EmployeeRequestsBannerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'EmployeeRequestsBannerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IEmployeeRequestsBannerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  cssurl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class EmployeeRequestsBannerApplicationCustomizer
  extends BaseApplicationCustomizer<IEmployeeRequestsBannerApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    const cssurl: string = this.properties.cssurl
    if (cssurl) {
      // inject the style sheet
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssurl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }



    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolders
    );

    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose }
        );
    }

    if (!this._topPlaceholder) {
      Log.error(
        LOG_SOURCE,
        new Error('The Top Placeholder is not available.')
      );
      return;
    }

    this._topPlaceholder.domElement.innerHTML = `
    <div style="
      background-color: #0078d4;
      color: white;
      padding: 12px 20px;
      text-align: center;
      font-size: 16px;
    ">
      Welcome to Employee Requests!
    </div>
  `;
  }

  private _onDispose(): void {
    console.log('Top Placeholder disposed.');
  }
}
