import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import * as strings from 'Monarch360NavCrudApplicationCustomizerStrings';
import { SettingsDialog } from './components/SettingsDialogNew';

const LOG_SOURCE: string = 'Monarch360NavCrud';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMonarch360NavCrudApplicationCustomizerProperties {
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class Monarch360NavCrudApplicationCustomizer
  extends BaseApplicationCustomizer<IMonarch360NavCrudApplicationCustomizerProperties> {

  private buttonInjectionInterval: number | null = null;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Try to inject immediately
    setTimeout(() => {
      this.injectSettingsButton();
    }, 500); // Initial delay to wait for SharePoint to render the header
    
    // Set up periodic check to ensure our button stays injected
    this.buttonInjectionInterval = window.setInterval(() => {
      const button = document.getElementById('monarch360SettingsBtn');
      const headerContainer = document.querySelector('[data-automationid="SiteHeader"]') || 
                              document.querySelector('[data-navigationcomponent="SiteHeader"]') ||
                              document.querySelector('.ms-siteHeader-container') ||
                              document.querySelector('#SuiteNavPlaceHolder');
      
      // Only inject if button is missing AND header container exists
      if (!button && headerContainer) {
        this.injectSettingsButton();
      } else if (button && !headerContainer) {
        // If the header container is gone but button still exists, remove the button
        button.remove();
      }
    }, 1000); // Check more frequently
    
    return Promise.resolve();
  }
  
  /**
   * Clean up event listeners on dispose
   */
  protected onDispose(): void {
    // Clear the interval
    if (this.buttonInjectionInterval) {
      window.clearInterval(this.buttonInjectionInterval);
      this.buttonInjectionInterval = null;
    }
    
    super.onDispose();
  }
  
  /**
   * Injects the settings button into the DOM
   */
  private injectSettingsButton = (): void => {
    try {
      // Remove any existing button first
      const existingButton = document.getElementById('monarch360SettingsBtn');
      if (existingButton) {
        existingButton.remove();
      }

      // Find the header container using multiple selectors for better compatibility
      const headerContainer = document.querySelector('[data-automationid="SiteHeader"]') || 
                               document.querySelector('[data-navigationcomponent="SiteHeader"]') ||
                               document.querySelector('.ms-siteHeader-container') ||
                               document.querySelector('#SuiteNavPlaceHolder');
      
      // Try to find a good insertion point within the header
      const insertionPoint = headerContainer?.querySelector('.ms-siteHeader-siteLogo') ||
                             headerContainer?.querySelector('[data-automationid="siteLogo"]') ||
                             headerContainer?.querySelector('.logoCell-110') ||
                             headerContainer?.querySelector('.ms-siteHeader-siteNav') ||
                             headerContainer?.firstElementChild;
      
      if (!headerContainer) {
        console.log('Header container not found, will retry later.');
        console.log('Available elements:', {
          'SiteHeader': document.querySelector('[data-automationid="SiteHeader"]'),
          'navigationcomponent': document.querySelector('[data-navigationcomponent="SiteHeader"]'),
          'siteHeader-container': document.querySelector('.ms-siteHeader-container'),
          'SuiteNavPlaceHolder': document.querySelector('#SuiteNavPlaceHolder')
        });
        return;
      }
      
      if (!insertionPoint) {
        console.log('Suitable insertion point not found, will retry later.');
        console.log('Header container found:', headerContainer);
        console.log('Available insertion points:', {
          'siteLogo': headerContainer?.querySelector('.ms-siteHeader-siteLogo'),
          'automationid-siteLogo': headerContainer?.querySelector('[data-automationid="siteLogo"]'),
          'logoCell': headerContainer?.querySelector('.logoCell-110'),
          'siteNav': headerContainer?.querySelector('.ms-siteHeader-siteNav'),
          'firstChild': headerContainer?.firstElementChild
        });
        return;
      }
      
      // Create settings button with gear icon
      const settingsBtn = document.createElement('button');
      settingsBtn.id = 'monarch360SettingsBtn';
      settingsBtn.title = 'Site Settings';
      settingsBtn.setAttribute('aria-label', 'Open site settings');
      
      // Apply styles - Updated to match SharePoint's modern UI
      settingsBtn.style.background = '#0078d4';
      settingsBtn.style.color = 'white';
      settingsBtn.style.border = 'none';
      settingsBtn.style.cursor = 'pointer';
      settingsBtn.style.padding = '8px';
      settingsBtn.style.marginRight = '10px';
      settingsBtn.style.display = 'flex';
      settingsBtn.style.alignItems = 'center';
      settingsBtn.style.justifyContent = 'center';
      settingsBtn.style.borderRadius = '4px';
      settingsBtn.style.height = '32px';
      settingsBtn.style.width = '32px';
      settingsBtn.style.zIndex = '1000';
      settingsBtn.style.position = 'relative';  // Added position
      settingsBtn.style.boxShadow = '0 2px 4px rgba(0, 0, 0, 0.1)';  // Add subtle shadow to match UI
      
      // Use Fluent UI gear icon
      settingsBtn.innerHTML = `<svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor">
        <path fill-rule="evenodd" clip-rule="evenodd" d="M11.49 3.17c-.38-1.56-2.6-1.56-2.98 0a1.532 1.532 0 01-2.286.948c-1.372-.836-2.942.734-2.106 2.106.54.886.29 2.045-.947 2.287-1.561.379-1.561 2.6 0 2.978a1.532 1.532 0 01.947 2.287c-.836 1.372.734 2.942 2.106 2.106a1.532 1.532 0 012.287.947c.379 1.561 2.6 1.561 2.978 0a1.533 1.533 0 012.287-.947c1.372.836 2.942-.734 2.106-2.106a1.533 1.533 0 01.947-2.287c1.561-.379 1.561-2.6 0-2.978a1.532 1.532 0 01-.947-2.287c.836-1.372-.734-2.942-2.106-2.106a1.532 1.532 0 01-2.287-.947zM10 13a3 3 0 100-6 3 3 0 000 6z"></path>
      </svg>`;
      
      // Add hover style and ensure visibility
      const style = document.createElement('style');
      style.innerHTML = `
        #monarch360SettingsBtn:hover {
          background-color: #106ebe !important;
        }
        #monarch360SettingsBtn:active {
          transform: scale(0.98);
        }
        /* Ensure button stays visible during page transitions */
        #monarch360SettingsBtn {
          visibility: visible !important;
          opacity: 1 !important;
          transition: background-color 0.2s ease;
          flex-shrink: 0; /* Prevent button from shrinking */
        }
        /* Improve visibility on various SharePoint themes */
        .ms-bgColor-themeDark #monarch360SettingsBtn,
        .ms-bgColor-neutralDark #monarch360SettingsBtn {
          background-color: #ffffff !important;
          color: #0078d4 !important;
        }
      `;
      document.head.appendChild(style);
      
      // Add click handler to show the settings dialog
      settingsBtn.onclick = () => {
        SettingsDialog.show(this.context);
      };
      
      // Insert button BEFORE the insertion point (to the left of the logo)
      if (insertionPoint.parentNode) {
        insertionPoint.parentNode.insertBefore(settingsBtn, insertionPoint);
      } else {
        // Fallback - prepend to the header container
        headerContainer.insertBefore(settingsBtn, headerContainer.firstChild);
      }
      
      // Apply any saved settings
      this.applyStoredSettings();
      
      console.log('Settings button successfully injected to the left of the site logo');
    } catch (error) {
      console.error('Error injecting settings button:', error);
      Log.error(LOG_SOURCE, error as Error);
    }
  }
  
  /**
   * Apply stored settings from SharePoint list on page load
   */
  private applyStoredSettings = (): void => {
    // Use a self-invoking async function since this method is called
    // in a synchronous context but needs to perform async operations
    (async () => {
      try {
        // Initialize SP instance
        const sp = spfi().using(SPFx(this.context));
        
        // Get settings from SharePoint list
        const bgColorItems = await sp.web.lists
          .getByTitle("navbarcrud")
          .items
          .filter("Title eq 'background_color'")
          .top(1)();
          
        const fontSizeItems = await sp.web.lists
          .getByTitle("navbarcrud")
          .items
          .filter("Title eq 'font_size'")
          .top(1)();
      
      // Check if we have settings items
      if (bgColorItems && bgColorItems.length > 0 && fontSizeItems && fontSizeItems.length > 0) {
        const backgroundColor = bgColorItems[0].value || '#0078d4'; // Default blue
        const fontSize = parseInt(fontSizeItems[0].value, 10) || 16; // Default size
        
        this.applySavedStyles(backgroundColor, fontSize);
        console.log('Applied settings from SharePoint list:', { backgroundColor, fontSize });
      } else {
        console.log('Settings not found in SharePoint list, using defaults');
        this.applySavedStyles('#0078d4', 16); // Apply defaults
      }
    } catch (error) {
      console.error('Error applying stored settings from SharePoint list:', error);
      
      // Check if the error is about missing list
      const errorMessage = (error as Error)?.message || '';
      if (errorMessage.indexOf('does not exist') !== -1 || errorMessage.indexOf('navbarcrud') !== -1) {
        console.warn('⚠️ SharePoint list "navbarcrud" not found. Please create it first.');
        console.log('📋 Required list structure:');
        console.log('  - List Name: navbarcrud');
        console.log('  - Columns: Title (text), value (text)');
        console.log('  - Items: background_color=#0078d4, font_size=16');
      }
      
      Log.error(LOG_SOURCE, error as Error);
      
      // Fallback to defaults on error
      this.applySavedStyles('#0078d4', 16);
    }
  })().catch((error) => {
    console.error('Error in applyStoredSettings:', error);
    this.applySavedStyles('#0078d4', 16);
  });
  }
  
  /**
   * Apply the saved styles to the page
   */
  private applySavedStyles = (color: string, fontSize: number): void => {
    // Calculate text color based on background brightness
    const getTextColor = (bgColor: string): string => {
      const r = parseInt(bgColor.substr(1, 2), 16);
      const g = parseInt(bgColor.substr(3, 2), 16);
      const b = parseInt(bgColor.substr(5, 2), 16);
      const brightness = (r * 299 + g * 587 + b * 114) / 1000;
      return brightness > 128 ? 'black' : 'white';
    };
    
    // Create a style element to inject our custom CSS
    const styleId = 'monarch360CustomStyles';
    let styleEl = document.getElementById(styleId) as HTMLStyleElement;
    
    // Create the style element if it doesn't exist
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = styleId;
      document.head.appendChild(styleEl);
    }
    
    // Apply styles specifically to the SharePoint navigation as required
    const cssRules = `
      /* SharePoint Header Background Color - Target specific spSiteHeader element */
      #spSiteHeader {
        background-color: ${color} !important;
      }
      
      /* Fallback selectors if spSiteHeader is not available */
      [data-automationid="ShyHeader"],
      [data-navigationcomponent="SiteHeader"],
      .ms-FocusZone.ms-siteHeader-siteNav {
        background-color: ${color} !important;
      }
      
      /* Target specific HorizontalNavItem elements */
      .ms-HorizontalNavItem[data-automationid="HorizontalNav-link"],
      .ms-HorizontalNavItem .ms-HorizontalNavItem-link,
      .ms-HorizontalNavItem .ms-HorizontalNavItem-linkText {
        font-size: ${fontSize}px !important;
        color: ${getTextColor(color)} !important;
      }
      
      /* Additional navigation elements font size */
      [data-automationid="ShyHeader"] .ms-HorizontalNavItem-linkText,
      [data-automationid="ShyHeader"] span,
      [data-automationid="ShyHeader"] a,
      [data-automationid="ShyHeader"] button,
      [data-navigationcomponent="SiteHeader"] span,
      [data-navigationcomponent="SiteHeader"] a,
      [data-navigationcomponent="SiteHeader"] button {
        font-size: ${fontSize}px !important;
      }
      
      /* Ensure content is visible against the background color */
      [data-automationid="ShyHeader"] .ms-HorizontalNavItem-linkText,
      [data-automationid="ShyHeader"] span,
      [data-automationid="ShyHeader"] a,
      [data-automationid="ShyHeader"] button,
      [data-navigationcomponent="SiteHeader"] span,
      [data-navigationcomponent="SiteHeader"] a,
      [data-navigationcomponent="SiteHeader"] button,
      .ms-FocusZone.ms-siteHeader-siteNav span,
      .ms-FocusZone.ms-siteHeader-siteNav a {
        color: ${getTextColor(color)} !important;
      }
      
      /* Specific styling for HorizontalNavItem links */
      .ms-HorizontalNavItem[role="listitem"] .ms-HorizontalNavItem-link {
        color: ${getTextColor(color)} !important;
        font-size: ${fontSize}px !important;
      }
      
      /* Hover states for navigation items */
      .ms-HorizontalNavItem:hover .ms-HorizontalNavItem-linkText,
      .ms-HorizontalNavItem:hover .ms-HorizontalNavItem-link {
        color: ${getTextColor(color)} !important;
        opacity: 0.8;
      }
      
      /* Make settings button always visible and match header color */
      #monarch360SettingsBtn {
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        background-color: ${color} !important;
        color: ${getTextColor(color)} !important;
      }
      
      /* Adjust the header logo size proportionally */
      [data-automationid="ShyHeader"] img,
      [data-navigationcomponent="SiteHeader"] img,
      .ms-siteHeader-siteLogo img {
        height: ${fontSize * 1.5}px !important;
      }
    `;
    
    // Update the style element content
    styleEl.textContent = cssRules;
    console.log(`Applied custom styles - Background Color to #spSiteHeader: ${color}, Font Size to .ms-HorizontalNavItem: ${fontSize}px`);
  }
}
