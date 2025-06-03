import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

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

/**
 * Shows a custom settings dialog with content
 */
class SettingsDialog {
  public static show(): void {
    const dialogContent: string = `
      <style>
        .settings-dialog {
          padding: 20px;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .settings-dialog h2 {
          margin-top: 0;
          font-size: 18px;
        }
        .settings-dialog .form-group {
          margin-bottom: 20px;
        }
        .settings-dialog label {
          display: block;
          margin-bottom: 5px;
          font-weight: 500;
        }
        .settings-dialog .color-picker-container {
          display: flex;
          align-items: center;
        }
        .settings-dialog .color-picker {
          width: 60px;
          height: 40px;
          margin-right: 10px;
        }
        .settings-dialog .color-preview {
          flex-grow: 1;
          height: 40px;
          border-radius: 4px;
          display: flex;
          align-items: center;
          justify-content: center;
          color: white;
          background-color: #0078d4;
          transition: background-color 0.2s;
        }
        .settings-dialog .size-slider-container {
          display: flex;
          align-items: center;
          margin-bottom: 10px;
        }
        .settings-dialog .size-slider {
          flex-grow: 1;
          margin-right: 10px;
        }
        .settings-dialog .font-preview {
          padding: 10px;
          border: 1px solid #ccc;
          border-radius: 4px;
          font-size: 16px;
          text-align: center;
        }
      </style>
      <div class="settings-dialog">
        <h2>Site Settings</h2>
        <div class="form-group">
          <label for="colorPicker">Background Color:</label>
          <div class="color-picker-container">
            <input type="color" id="colorPicker" class="color-picker" value="#0078d4">
            <div id="colorPreview" class="color-preview">Preview</div>
          </div>
        </div>
        <div class="form-group">
          <label for="sizeSlider">Font Size:</label>
          <div class="size-slider-container">
            <input type="range" id="sizeSlider" class="size-slider" min="12" max="24" value="16">
            <span id="sizeValue">16px</span>
          </div>
          <div id="fontSizePreview" class="font-preview">Font Size Preview Text</div>
        </div>
        <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 15px;">
          <button id="settingsPrevBtn" style="padding: 8px 16px; background-color: #f3f2f1; color: #333; border: 1px solid #d2d0ce; border-radius: 2px; cursor: pointer;">
            Previous
          </button>
          <button id="settingsSaveBtn" style="padding: 8px 16px; background-color: #0078d4; color: white; border: none; border-radius: 2px; cursor: pointer;">
            Save Changes
          </button>
        </div>
      </div>
    `;

    // Show the dialog
    const dialog = Dialog.alert(dialogContent);

    // Use the dialog object to manipulate the dialog if needed
    // const dialogInstance = dialog;

    // Wait for the DOM to be updated with the dialog content
    setTimeout(() => {
      // Get DOM elements
      const colorPicker = document.getElementById('colorPicker');
      const colorPreview = document.getElementById('colorPreview');
      const sizeSlider = document.getElementById('sizeSlider');
      const sizeValue = document.getElementById('sizeValue');
      const fontSizePreview = document.getElementById('fontSizePreview');
      const saveBtn = document.getElementById('settingsSaveBtn');
      const prevBtn = document.getElementById('settingsPrevBtn');
      
      if (colorPicker instanceof HTMLInputElement && colorPreview) {
        colorPicker.addEventListener('input', () => {
          const selectedColor = colorPicker.value;
          colorPreview.style.backgroundColor = selectedColor;
          
          // Determine if text should be white or black based on color brightness
          const r = parseInt(selectedColor.substr(1, 2), 16);
          const g = parseInt(selectedColor.substr(3, 2), 16);
          const b = parseInt(selectedColor.substr(5, 2), 16);
          const brightness = (r * 299 + g * 587 + b * 114) / 1000;
          colorPreview.style.color = brightness > 128 ? 'black' : 'white';
        });
      }
      
      if (sizeSlider instanceof HTMLInputElement && sizeValue && fontSizePreview) {
        sizeSlider.addEventListener('input', () => {
          const val = sizeSlider.value;
          sizeValue.textContent = `${val}px`;
          fontSizePreview.style.fontSize = `${val}px`;
        });
      }
      
      if (prevBtn) {
        prevBtn.addEventListener('click', () => {
          console.log('Previous button clicked');
          // Add your previous step logic here
          alert('Going back to previous step');
        });
      }
      
      if (saveBtn) {
        saveBtn.addEventListener('click', () => {
          const color = colorPicker instanceof HTMLInputElement ? colorPicker.value : '#0078d4';
          const size = sizeSlider instanceof HTMLInputElement ? parseInt(sizeSlider.value, 10) : 16;
          
          console.log('Settings saved:', { color, size });
          
          // Close dialog using the native dialog close button
          const closeButton = document.querySelector('.ms-Dialog-button');
          if (closeButton instanceof HTMLElement) {
            closeButton.click();
          }
        });
      }
    }, 100);
  }
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
          SettingsDialog.show();
        };
        // Insert as first child
        logoCell.insertBefore(settingsBtn, logoCell.firstChild);
      }
    }, 1000); // Increased timeout for more reliable loading

    return Promise.resolve();
  }
}
