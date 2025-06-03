import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'Monarch360NavCrudApplicationCustomizerStrings';

const LOG_SOURCE: string = 'Monarch360NavCrudApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMonarch360NavCrudApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class Monarch360NavCrudApplicationCustomizer
  extends BaseApplicationCustomizer<IMonarch360NavCrudApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Wait for DOM and inject settings icon
    setTimeout(() => {
      // Using data attributes for more reliable selection
      const logoLink = document.querySelector('a[data-navigationcomponent="SiteHeader"][data-interception="propagate"]');
      if (logoLink && logoLink.parentElement) {
        const logoCell = logoLink.parentElement;
        // Create settings icon button
        const settingsBtn = document.createElement('button');
        settingsBtn.title = 'Settings';
        settingsBtn.style.background = 'none';
        settingsBtn.style.border = 'none';
        settingsBtn.style.cursor = 'pointer';
        settingsBtn.style.marginRight = '8px';
        settingsBtn.innerHTML = `<svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
        <path fill-rule="evenodd" clip-rule="evenodd" d="M11.49 3.17c-.38-1.56-2.6-1.56-2.98 0a1.532 1.532 0 01-2.286.948c-1.372-.836-2.942.734-2.106 2.106.54.886.29 2.045-.947 2.287-1.561.379-1.561 2.6 0 2.978a1.532 1.532 0 01.947 2.287c-.836 1.372.734 2.942 2.106 2.106a1.532 1.532 0 012.287.947c.379 1.561 2.6 1.561 2.978 0a1.533 1.533 0 012.287-.947c1.372.836 2.942-.734 2.106-2.106a1.533 1.533 0 01.947-2.287c1.561-.379 1.561-2.6 0-2.978a1.532 1.532 0 01-.947-2.287c.836-1.372-.734-2.942-2.106-2.106a1.532 1.532 0 01-2.287-.947zM10 13a3 3 0 100-6 3 3 0 000 6z" />
        </svg>`;
        // Optional: Add click handler
        settingsBtn.onclick = () => {
          alert('Settings clicked!');
        };
        // Insert as first child
        logoCell.insertBefore(settingsBtn, logoCell.firstChild);
      }
    }, 1000); // Increased timeout for more reliable loading

    return Promise.resolve();
  }
}
