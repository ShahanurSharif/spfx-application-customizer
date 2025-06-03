import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'Monarch360NavCrudApplicationCustomizerStrings';
import { SettingsDialog } from './components/SettingsDialog';

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
    this.injectSettingsButton();

    // Listen for navigation events to handle SPA navigation
    window.addEventListener('hashchange', this.injectSettingsButton);
    
    return Promise.resolve();
  }
  
  /**
   * Clean up event listeners on dispose
   */
  protected onDispose(): void {
    window.removeEventListener('hashchange', this.injectSettingsButton);
    super.onDispose();
  }
  
  /**
   * Injects the settings button into the DOM
   */
  private injectSettingsButton = (): void => {
    // Ensure previous button is removed if it exists
    const existingButton = document.getElementById('monarch360SettingsBtn');
    if (existingButton) {
      existingButton.remove();
    }
    
    setTimeout(() => {
      try {
        // Using data attributes for more reliable selection
        const logoLink = document.querySelector('a[data-navigationcomponent="SiteHeader"][data-interception="propagate"]');
        if (logoLink && logoLink.parentElement) {
          const logoCell = logoLink.parentElement;
          
          // Create settings icon button with improved styling
          const settingsBtn = document.createElement('button');
          settingsBtn.id = 'monarch360SettingsBtn';
          settingsBtn.title = 'Site Settings';
          settingsBtn.setAttribute('aria-label', 'Open site settings');
          
          // Apply styles
          settingsBtn.style.background = 'none';
          settingsBtn.style.border = 'none';
          settingsBtn.style.cursor = 'pointer';
          settingsBtn.style.marginRight = '10px';
          settingsBtn.style.padding = '8px';
          settingsBtn.style.display = 'flex';
          settingsBtn.style.alignItems = 'center';
          settingsBtn.style.justifyContent = 'center';
          settingsBtn.style.borderRadius = '4px';
          settingsBtn.style.transition = 'background-color 0.2s';
          
          // Hover effect using CSS
          const style = document.createElement('style');
          style.innerHTML = `
            #monarch360SettingsBtn:hover {
              background-color: rgba(0, 0, 0, 0.04);
            }
            #monarch360SettingsBtn:active {
              background-color: rgba(0, 0, 0, 0.08);
            }
            #monarch360SettingsBtn svg {
              transition: transform 0.3s ease;
            }
            #monarch360SettingsBtn:hover svg {
              transform: rotate(30deg);
            }
          `;
          document.head.appendChild(style);
          
          // Use Fluent UI gear icon
          settingsBtn.innerHTML = `<svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
            <path fill-rule="evenodd" clip-rule="evenodd" d="M11.49 3.17c-.38-1.56-2.6-1.56-2.98 0a1.532 1.532 0 01-2.286.948c-1.372-.836-2.942.734-2.106 2.106.54.886.29 2.045-.947 2.287-1.561.379-1.561 2.6 0 2.978a1.532 1.532 0 01.947 2.287c-.836 1.372.734 2.942 2.106 2.106a1.532 1.532 0 012.287.947c.379 1.561 2.6 1.561 2.978 0a1.533 1.533 0 012.287-.947c1.372.836 2.942-.734 2.106-2.106a1.533 1.533 0 01.947-2.287c1.561-.379 1.561-2.6 0-2.978a1.532 1.532 0 01-.947-2.287c.836-1.372-.734-2.942-2.106-2.106a1.532 1.532 0 01-2.287-.947zM10 13a3 3 0 100-6 3 3 0 000 6z" />
          </svg>`;
          
          // Add click handler to show the settings dialog
          settingsBtn.onclick = () => {
            SettingsDialog.show();
          };
          
          // Insert as first child
          logoCell.insertBefore(settingsBtn, logoCell.firstChild);
          
          // Apply any saved settings on page load
          this.applyStoredSettings();
        }
      } catch (error) {
        Log.error(LOG_SOURCE, error);
      }
    }, 1000); // Increased timeout for more reliable loading
  }
  
  /**
   * Apply stored settings on page load
   */
  private applyStoredSettings = (): void => {
    try {
      const savedSettings = localStorage.getItem('monarch360Settings');
      
      if (savedSettings) {
        const settings = JSON.parse(savedSettings);
        
        if (settings.backgroundColor) {
          const suiteBar = document.querySelector('.ms-CommandBar');
          if (suiteBar) {
            suiteBar.setAttribute('style', `background-color: ${settings.backgroundColor} !important`);
          }
        }
        
        if (settings.fontSize) {
          document.body.style.setProperty('--default-font-size', `${settings.fontSize}px`);
        }
      }
    } catch (error) {
      const errorMessage = error instanceof Error ? error : new Error(String(error));
      Log.error(LOG_SOURCE, errorMessage);
    }
  }
}
