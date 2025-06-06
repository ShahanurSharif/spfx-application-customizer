import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { 
  Dialog, 
  DialogType, 
  IDialogContentProps, 
  DialogFooter,
  PrimaryButton, 
  Label,
  Slider,
  TextField,
  Stack,
  StackItem,
  mergeStyles,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  IStackTokens
} from '@fluentui/react';

// Import PnP JS libraries
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';

// Context for SharePoint operations
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

/**
 * Settings Interface
 */
export interface IUserSettings {
  backgroundColor: string;
  fontSize: number;
  logoFile?: File; // Changed from logo URL to logo File
  logoUrl?: string; // Keep URL for displaying current logo
}

/**
 * Props for the Settings Dialog Component
 */
export interface ISettingsDialogProps {
  onDismiss: () => void;
  context?: ApplicationCustomizerContext; // SharePoint context
}

/**
 * State for the Settings Dialog Component
 */
interface ISettingsDialogState {
  hideDialog: boolean;
  color: string;
  fontSize: number;
  logoFile: File | undefined; // Changed to File object
  logoUrl: string; // URL for displaying current logo
  logoPreview: string; // Preview URL for uploaded file
  previewText: string;
  isLoading: boolean;
  isSaving: boolean;
  notification: {
    message: string;
    type: MessageBarType;
    show: boolean;
  };
}

/**
 * Settings Dialog Component
 */
export class SettingsDialogContent extends React.Component<ISettingsDialogProps, ISettingsDialogState> {
  constructor(props: ISettingsDialogProps) {
    super(props);
    
    this.state = {
      hideDialog: false,
      color: '#0078d4',
      fontSize: 16,
      logoFile: undefined, // Initialize as undefined
      logoUrl: '', // Initialize empty logo URL
      logoPreview: '', // Initialize empty logo preview
      previewText: 'Settings Preview',
      isLoading: true,
      isSaving: false,
      notification: {
        message: '',
        type: MessageBarType.info,
        show: false
      }
    };
  }

  /**
   * Component Did Mount - Load settings
   */
  public componentDidMount(): void {
    // Use a self-invoking async function to handle the promise
    (async () => {
      try {
        await this.loadSettings();
        // After loading settings, force logo replacement if we have a logo
        await this.forceLogoReplacementOnLoad();
      } catch (error) {
        console.error('Error loading settings:', error);
        this.setState({ 
          isLoading: false,
          notification: {
            message: 'Error loading settings',
            type: MessageBarType.error,
            show: true
          }
        });
      }
    })().catch((error) => {
      console.error('Error in componentDidMount:', error);
      this.setState({ 
        isLoading: false,
        notification: {
          message: 'Error loading settings',
            type: MessageBarType.error,
            show: true
          }
        });
    });
  }
  
  /**
   * Load settings from SharePoint list
   */
  private loadSettings = async (): Promise<void> => {
    try {
      if (!this.props.context) {
        console.error('SharePoint context not available');
        this.setState({ 
          isLoading: false,
          notification: {
            message: 'SharePoint context not available',
            type: MessageBarType.error,
            show: true
          }
        });
        return;
      }
      
      console.log('Loading settings from SharePoint list...');
      
      // Initialize SP
      const sp = spfi().using(SPFx(this.props.context));
      
      // Get background color from list
      console.log('Fetching background color...');
      const bgColorItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'background_color'")
        .select("value")
        .top(1)();
        
      // Get font size from list
      console.log('Fetching font size...');
      const fontSizeItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'font_size'")
        .select("value")
        .top(1)();
        
      // Get logo from list
      console.log('Fetching logo...');
      const logoItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'logo'")
        .select("value")
        .top(1)();
        
      // Default values
      let backgroundColor = '#0078d4';
      let fontSize = 16;
      let logoUrl = '';
      
      // If items exist, use those values
      if (bgColorItems.length > 0) {
        backgroundColor = bgColorItems[0].value || backgroundColor;
      }
      
      if (fontSizeItems.length > 0) {
        fontSize = parseInt(fontSizeItems[0].value, 10) || fontSize;
      }
      
      if (logoItems.length > 0 && logoItems[0].value) {
        // Handle logo URL from value field, same as other settings
        logoUrl = logoItems[0].value || logoUrl;
      }
      
      this.setState({
        color: backgroundColor,
        fontSize: fontSize,
        logoUrl: logoUrl,
        isLoading: false,
        notification: {
          message: 'Settings loaded successfully from SharePoint list',
          type: MessageBarType.success,
          show: true
        }
      });
      
      // Apply the loaded settings to the page immediately
      this.applySettings(backgroundColor, fontSize, logoUrl);
      
      // Hide notification after 3 seconds
      setTimeout(() => {
        this.setState({ 
          notification: { 
            ...this.state.notification,
            show: false 
          } 
        });
      }, 3000);
      
    } catch (error) {
      console.error('Error loading settings from SharePoint list:', error);
      this.setState({ 
        isLoading: false,
        notification: {
          message: 'Error loading settings from SharePoint',
          type: MessageBarType.error,
          show: true
        }
      });
    }
  }
  
  /**
   * Save settings to SharePoint list
   */
  private saveSettings = async (settings: IUserSettings): Promise<void> => {
    try {
      if (!this.props.context) {
        console.error('SharePoint context not available');
        throw new Error('SharePoint context not available');
      }
      
      // Initialize SP
      const sp = spfi().using(SPFx(this.props.context));
      
      // Check if the background_color item exists
      const bgColorItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'background_color'")
        .top(1)();
        
      if (bgColorItems.length > 0) {
        // Update the existing item
        await sp.web.lists.getByTitle("navbarcrud").items
          .getById(bgColorItems[0].Id)
          .update({
            value: settings.backgroundColor
          });
      } else {
        // Create a new item
        await sp.web.lists.getByTitle("navbarcrud").items
          .add({
            Title: 'background_color',
            value: settings.backgroundColor
          });
      }
      
      // Check if the font_size item exists
      const fontSizeItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'font_size'")
        .top(1)();
        
      if (fontSizeItems.length > 0) {
        // Update the existing item
        await sp.web.lists.getByTitle("navbarcrud").items
          .getById(fontSizeItems[0].Id)
          .update({
            value: settings.fontSize.toString()
          });
      } else {
        // Create a new item
        await sp.web.lists.getByTitle("navbarcrud").items
          .add({
            Title: 'font_size',
            value: settings.fontSize.toString()
          });
      }
      
      // Handle logo file upload if provided
      let logoUrl = '';
      if (settings.logoFile) {
        try {
          console.log('Starting logo file upload...', settings.logoFile.name);
          
          // Upload file to SharePoint Site Assets
          const fileName = `site-logo-${Date.now()}.${settings.logoFile.name.split('.').pop()}`;
          const fileContent = await this.fileToArrayBuffer(settings.logoFile);
          
          console.log('Uploading file to Site Assets:', fileName);
          
          // Upload to Site Assets library
          const fileAddResult = await sp.web.lists.getByTitle("Site Assets").rootFolder.files
            .addUsingPath(fileName, fileContent, { Overwrite: true });
          
          // Get the server relative URL from the file info
          logoUrl = fileAddResult.ServerRelativeUrl;
          console.log('Logo uploaded successfully:', logoUrl);
        } catch (uploadError) {
          console.error('Error uploading logo file:', uploadError);
          throw new Error('Failed to upload logo file');
        }
      } else if (settings.logoUrl) {
        // Use existing logo URL
        logoUrl = settings.logoUrl;
      }
      
      // Check if the logo item exists
      console.log('Checking for existing logo item...');
      
      const logoItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'logo'")
        .top(1)();
        
      console.log('Logo items found:', logoItems.length);
      console.log('Logo URL to save:', logoUrl);
        
      if (logoItems.length > 0) {
        // Update the existing item - Use 'value' field like other settings
        console.log('Updating existing logo item...');
        await sp.web.lists.getByTitle("navbarcrud").items
          .getById(logoItems[0].Id)
          .update({
            value: logoUrl || ''
          });
        console.log('Logo item updated successfully');
      } else {
        // Create a new item - Use 'value' field like other settings
        console.log('Creating new logo item...');
        await sp.web.lists.getByTitle("navbarcrud").items
          .add({
            Title: 'logo',
            value: logoUrl || ''
          });
        console.log('Logo item created successfully');
      }
      
      // Update component state with new logo URL if file was uploaded
      if (settings.logoFile && logoUrl) {
        this.setState({ logoUrl: logoUrl });
      }
      
      console.log('Settings saved to SharePoint list successfully!');
    } catch (error) {
      console.error('Error saving settings to SharePoint list:', error);
      throw error;
    }
  }

  /**
   * Handles dialog close
   */
  private closeDialog = (): void => {
    this.setState({ hideDialog: true });
    this.props.onDismiss();
  }

  /**
   * Applies the settings to the page
   */
  private applySettings = (color: string, fontSize: number, logo?: string): void => {
    try {
      // Create a style element to inject our custom CSS
      const styleId = 'monarch360CustomStyles';
      let styleEl = document.getElementById(styleId) as HTMLStyleElement;
      
      // Create the style element if it doesn't exist
      if (!styleEl) {
        styleEl = document.createElement('style');
        styleEl.id = styleId;
        document.head.appendChild(styleEl);
      }
      
      // Apply styles specifically to ShyHeader as required
      const cssRules = `
        /* Shy Header Background Color */
        [data-automationid="ShyHeader"] {
          background-color: ${color} !important;
        }
        
        /* Also target SuiteNav for complete header styling */
        .ms-FocusZone.ms-siteHeader-siteNav {
          background-color: ${color} !important;
        }
        
        /* Adjust font size of navigation links */
        [data-automationid="ShyHeader"] .ms-HorizontalNavItem-linkText,
        [data-automationid="ShyHeader"] span,
        [data-automationid="ShyHeader"] a,
        [data-automationid="ShyHeader"] button {
          font-size: ${fontSize}px !important;
        }
        
        /* Ensure content is visible against the background color */
        [data-automationid="ShyHeader"] .ms-HorizontalNavItem-linkText,
        [data-automationid="ShyHeader"] span,
        [data-automationid="ShyHeader"] a,
        [data-automationid="ShyHeader"] button,
        .ms-FocusZone.ms-siteHeader-siteNav span,
        .ms-FocusZone.ms-siteHeader-siteNav a {
          color: ${this.getTextColor(color)} !important;
        }
        
        /* Adjust the header logo size proportionally */
        [data-automationid="ShyHeader"] img,
        .ms-siteHeader-siteLogo img {
          height: ${fontSize * 1.5}px !important;
        }
        
        /* Make settings button always visible */
        #monarch360SettingsBtn {
          background-color: ${color} !important;
          color: ${this.getTextColor(color)} !important;
        }
      `;
      
      // Update the style element content
      styleEl.textContent = cssRules;
      
      // Apply logo if provided - use more aggressive and persistent approach
      if (logo && logo.trim()) {
        console.log('🎯 Applying logo with aggressive strategy:', logo);
        
        // Apply immediately with multiple attempts
        this.applySiteLogo(logo);
        
        // Apply again after short delays to catch dynamically loaded elements
        setTimeout(() => this.applySiteLogo(logo), 100);
        setTimeout(() => this.applySiteLogo(logo), 300);
        setTimeout(() => this.applySiteLogo(logo), 500);
        setTimeout(() => this.applySiteLogo(logo), 1000);
        
        // Set up the aggressive observer
        this.setupLogoObserver(logo);
        
        // Also inject CSS to override SharePoint's logo loading
        const logoOverrideCSS = `
          /* Override SharePoint default logos */
          img[src*="siteiconmanager"] {
            content: url("${logo}") !important;
          }
          img[src*="_api/siteiconmanager"] {
            content: url("${logo}") !important;
          }
          .logoImg-112 {
            content: url("${logo}") !important;
          }
          [class*="logoImg"] img {
            content: url("${logo}") !important;
          }
        `;
        
        // Inject the logo override CSS
        const logoStyleId = 'monarch360LogoOverride';
        let logoStyleEl = document.getElementById(logoStyleId) as HTMLStyleElement;
        
        if (!logoStyleEl) {
          logoStyleEl = document.createElement('style');
          logoStyleEl.id = logoStyleId;
          document.head.appendChild(logoStyleEl);
        }
        
        logoStyleEl.textContent = logoOverrideCSS;
        
        console.log('✅ Logo CSS override injected');
      }
      
      console.log(`Applied settings - Color: ${color}, Font Size: ${fontSize}px, Logo: ${logo || 'none'}`);
    } catch (error) {
      console.error('Failed to apply settings:', error);
    }
  }
  
  /**
   * Applies the site logo with aggressive replacement strategy
   */
  private applySiteLogo = (logoUrl: string): void => {
    try {
      if (!logoUrl || logoUrl.trim() === '') {
        console.log('No logo URL provided, skipping logo application');
        return;
      }

      console.log('🎯 Starting aggressive logo replacement with URL:', logoUrl);

      // Enhanced logo selectors targeting SharePoint's default logo API and elements
      const logoSelectors = [
        // Standard SharePoint logo selectors
        '[data-automationid="ShyHeader"] img',
        '.ms-siteHeader-siteLogo img',
        '.ms-siteLogo-img',
        '[data-automation-id="siteLogo"] img',
        
        // Specific selectors for SharePoint's default logo API
        'img[src*="siteiconmanager/getsitelogo"]',
        'img[src*="_api/siteiconmanager"]',
        'img[src*="/_api/siteiconmanager/getsitelogo"]',
        
        // Class-based selectors from HTML structure
        '.logoImg-112',
        '[class*="logoImg"]',
        'img[class*="logoImg"]',
        
        // Broader selectors for any missed logos
        '[data-automationid="ShyHeader"] [class*="logo"] img',
        '.ms-siteHeader [class*="logo"] img',
        'header img[alt*="logo"], header img[alt*="Logo"]'
      ];
      
      let logoApplied = false;
      let totalElementsFound = 0;

      logoSelectors.forEach(selector => {
        try {
          const logoElements = document.querySelectorAll(selector);
          totalElementsFound += logoElements.length;
          
          logoElements.forEach((logoElement: Element) => {
            if (logoElement && logoElement.tagName === 'IMG') {
              const imgElement = logoElement as HTMLImageElement;
              
              // Check if this is SharePoint's default logo or needs replacement
              const shouldReplace = !imgElement.src.includes(logoUrl) && (
                imgElement.src.includes('siteiconmanager') ||
                imgElement.src.includes('_api/siteiconmanager') ||
                imgElement.className.includes('logoImg') ||
                selector.includes('logoImg')
              );
              
              if (shouldReplace || imgElement.src.includes('siteiconmanager')) {
                console.log(`🔄 Replacing logo with selector: ${selector}`, {
                  oldSrc: imgElement.src,
                  newSrc: logoUrl,
                  element: imgElement
                });
                
                // Apply the new logo
                imgElement.src = logoUrl;
                imgElement.alt = 'Custom Site Logo';
                
                // Force visibility and styling
                imgElement.style.display = 'block !important';
                imgElement.style.visibility = 'visible !important';
                imgElement.style.opacity = '1';
                imgElement.style.maxHeight = '40px';
                imgElement.style.width = 'auto';
                
                // Trigger a re-render by temporarily hiding and showing
                imgElement.style.opacity = '0';
                setTimeout(() => {
                  imgElement.style.opacity = '1';
                }, 10);
                
                logoApplied = true;
              }
            }
          });
        } catch (selectorError) {
          console.warn(`Error with selector ${selector}:`, selectorError);
        }
      });
      
      console.log(`📊 Logo replacement summary:`, {
        logoApplied,
        totalElementsFound,
        logoUrl,
        allImages: document.querySelectorAll('img').length
      });
      
      if (logoApplied) {
        console.log(`✅ Applied site logo successfully: ${logoUrl}`);
      } else {
        console.warn('⚠️ No logo elements found to replace. Debugging info:', {
          totalImages: document.querySelectorAll('img').length,
          sharepointLogos: document.querySelectorAll('img[src*="siteiconmanager"]').length,
          logoImgElements: document.querySelectorAll('[class*="logoImg"]').length
        });
      }
    } catch (error) {
      console.error('Failed to apply site logo:', error);
    }
  }
  
  /**
   * Sets up an aggressive MutationObserver to watch for dynamically loaded SharePoint logos
   */
  private setupLogoObserver = (logoUrl: string): void => {
    try {
      console.log('👀 Setting up aggressive logo observer for:', logoUrl);
      
      // Create a MutationObserver to watch for changes to the logo
      const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
          if (mutation.type === 'childList' || mutation.type === 'attributes') {
            
            // Check for new image elements or src changes
            const checkAndReplaceLogos = (): void => {
              // Enhanced selectors for aggressive logo detection
              const logoSelectors = [
                'img[src*="siteiconmanager"]',
                'img[src*="_api/siteiconmanager"]',
                'img[src*="getsitelogo"]',
                '.logoImg-112',
                '[class*="logoImg"]',
                'img[class*="logoImg"]',
                '[data-automationid="ShyHeader"] img',
                '.ms-siteHeader img'
              ];
              
              logoSelectors.forEach(selector => {
                const logoElements = document.querySelectorAll(selector);
                logoElements.forEach((logoElement: Element) => {
                  if (logoElement && logoElement.tagName === 'IMG') {
                    const imgElement = logoElement as HTMLImageElement;
                    
                    // Check if this is a SharePoint default logo that needs replacement
                    const isSharePointLogo = imgElement.src.includes('siteiconmanager') ||
                                           imgElement.src.includes('_api/siteiconmanager') ||
                                           imgElement.src.includes('getsitelogo') ||
                                           imgElement.className.includes('logoImg');
                    
                    const isNotCustomLogo = !imgElement.src.includes(logoUrl);
                    
                    if (isSharePointLogo && isNotCustomLogo) {
                      console.log('🔄 Observer detected SharePoint logo, replacing:', {
                        selector,
                        oldSrc: imgElement.src,
                        newSrc: logoUrl,
                        element: imgElement
                      });
                      
                      // Replace with custom logo
                      imgElement.src = logoUrl;
                      imgElement.alt = 'Custom Site Logo';
                      
                      // Force styling
                      imgElement.style.display = 'block !important';
                      imgElement.style.visibility = 'visible !important';
                      imgElement.style.opacity = '1';
                      
                      // Trigger re-render
                      imgElement.style.opacity = '0';
                      setTimeout(() => {
                        imgElement.style.opacity = '1';
                      }, 10);
                    }
                  }
                });
              });
            };
            
            // Check immediately
            checkAndReplaceLogos();
            
            // Also check after a short delay to catch async loaded images
            setTimeout(checkAndReplaceLogos, 100);
          }
        });
      });
      
      // Start observing the document for changes
      observer.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['src', 'class', 'style']
      });
      
      // Also set up periodic checks for the first few seconds
      const intervals: NodeJS.Timeout[] = [];
      
      // Check every 500ms for the first 5 seconds
      for (let i = 1; i <= 10; i++) {
        const interval = setTimeout(() => {
          console.log(`🔍 Periodic logo check #${i}`);
          this.applySiteLogo(logoUrl);
        }, i * 500);
        intervals.push(interval);
      }
      
      // Stop observing after 15 seconds and clear intervals
      setTimeout(() => {
        observer.disconnect();
        intervals.forEach(interval => clearTimeout(interval));
        console.log('🛑 Logo observer and periodic checks stopped');
      }, 15000);
      
      console.log('👀 Logo observer started for 15 seconds with periodic checks');
    } catch (error) {
      console.error('Error setting up logo observer:', error);
    }
  }

  /**
   * Forces logo replacement on page load - ensures custom logo overrides SharePoint's default
   */
  private forceLogoReplacementOnLoad = async (): Promise<void> => {
    try {
      // Wait for potential logo to be loaded from SharePoint list
      if (this.state.logoUrl && this.state.logoUrl.trim()) {
        console.log('🚀 Forcing logo replacement on load:', this.state.logoUrl);
        
        // Apply logo multiple times with increasing delays to catch SharePoint's async loading
        const delays = [0, 100, 300, 500, 1000, 2000, 3000];
        
        delays.forEach(delay => {
          setTimeout(() => {
            this.applySiteLogo(this.state.logoUrl);
            console.log(`🔄 Logo applied after ${delay}ms delay`);
          }, delay);
        });
        
        // Set up the observer for continuous monitoring
        this.setupLogoObserver(this.state.logoUrl);
      }
    } catch (error) {
      console.error('Error in forceLogoReplacementOnLoad:', error);
    }
  }

  /**
   * Shows a notification
   */
  private showNotification = (message: string, type: MessageBarType): void => {
    this.setState({
      notification: {
        message,
        type,
        show: true
      }
    });
    
    // Auto-hide after 3 seconds
    setTimeout(() => {
      this.setState({ 
        notification: { 
          ...this.state.notification,
          show: false 
        } 
      });
    }, 3000);
  }
  
  /**
   * Handles save button click
   */
  private handleSaveClick = async (): Promise<void> => {
    const { color, fontSize, logoFile } = this.state;
    
    try {
      // Set saving state
      this.setState({ isSaving: true });
      
      // Save settings to SharePoint list
      const settings: IUserSettings = {
        backgroundColor: color,
        fontSize: fontSize,
        logoFile: logoFile
      };
      
      await this.saveSettings(settings);
      
      // Apply the settings with the logo URL (after upload if file was provided)
      this.applySettings(color, fontSize, this.state.logoUrl);
      
      // Show success notification
      this.showNotification('Settings saved successfully! Reloading page...', MessageBarType.success);
      
      // Clear saving state
      this.setState({ isSaving: false });
      
      // Close dialog after a short delay, then reload the page
      setTimeout(() => {
        this.closeDialog();
        // Reload the page to ensure all settings are properly applied
        window.location.reload();
      }, 1500);
    } catch (error) {
      console.error('Error saving settings to SharePoint list:', error);
      this.setState({ isSaving: false });
      this.showNotification('Error saving settings to SharePoint list', MessageBarType.error);
    }
  }

  /**
   * Handles color change
   */
  private handleColorChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const color = event.target.value;
    this.setState({ color });
  }

  /**
   * Handles font size change
   */
  private handleFontSizeChange = (value: number): void => {
    this.setState({ fontSize: value });
  }

  /**
   * Handles logo file change
   */
  private handleLogoChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const file = files[0];
      
      // Validate file type
      if (!file.type.startsWith('image/')) {
        this.showNotification('Please select a valid image file.', MessageBarType.error);
        return;
      }
      
      // Validate file size (max 5MB)
      const maxSize = 5 * 1024 * 1024; // 5MB in bytes
      if (file.size > maxSize) {
        this.showNotification('File size must be less than 5MB.', MessageBarType.error);
        return;
      }
      
      // Create preview URL
      const previewUrl = URL.createObjectURL(file);
      
      this.setState({ 
        logoFile: file,
        logoPreview: previewUrl
      });
    }
  }

  /**
   * Convert File to ArrayBuffer for SharePoint upload
   */
  private fileToArrayBuffer = (file: File): Promise<ArrayBuffer> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (reader.result instanceof ArrayBuffer) {
          resolve(reader.result);
        } else {
          reject(new Error('Failed to read file as ArrayBuffer'));
        }
      };
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Calculates text color based on background brightness
   */
  private getTextColor(bgColor: string): string {
    const r = parseInt(bgColor.substr(1, 2), 16);
    const g = parseInt(bgColor.substr(3, 2), 16);
    const b = parseInt(bgColor.substr(5, 2), 16);
    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
    return brightness > 128 ? 'black' : 'white';
  }

  /**
   * Component Will Unmount - Cleanup
   */
  public componentWillUnmount(): void {
    // Clean up object URL to prevent memory leaks
    if (this.state.logoPreview) {
      URL.revokeObjectURL(this.state.logoPreview);
    }
  }
  
  /**
   * Renders the component
   */
  public render(): React.ReactElement<ISettingsDialogProps> {
    const { hideDialog, color, fontSize, isLoading, isSaving, notification } = this.state;
    
    const dialogContentProps: IDialogContentProps = {
      type: DialogType.normal,
      title: 'ShyHeader Settings',
      closeButtonAriaLabel: 'Close'
    };

    const stackTokens: IStackTokens = { childrenGap: 16 };

    const combinedPreviewClass = mergeStyles({
      backgroundColor: color,
      color: this.getTextColor(color),
      padding: '20px',
      borderRadius: '4px',
      fontSize: `${fontSize}px`,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      marginTop: '16px',
      border: '1px solid #ccc',
      boxShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
      minHeight: '80px',
      textAlign: 'center'
    });

    return (
      <Dialog
        hidden={hideDialog}
        onDismiss={this.closeDialog}
        dialogContentProps={dialogContentProps}
        minWidth={400}
        maxWidth="90%"
      >
        {notification.show && (
          <MessageBar
            messageBarType={notification.type}
            isMultiline={false}
            onDismiss={() => this.setState({ notification: { ...notification, show: false } })}
            dismissButtonAriaLabel="Close"
            styles={{ root: { marginBottom: 15 } }}
          >
            {notification.message}
          </MessageBar>
        )}

        {isLoading ? (
          <Stack tokens={stackTokens} horizontalAlign="center" styles={{ root: { padding: '20px 0' } }}>
            <Spinner size={SpinnerSize.large} label="Loading settings..." />
          </Stack>
        ) : (
          <>
            <Stack tokens={stackTokens}>
              <StackItem>
                <Label>ShyHeader Background Color:</Label>
                <Stack horizontal tokens={{ childrenGap: 10 }}>
                  <StackItem>
                    <input 
                      type="color" 
                      value={color} 
                      onChange={this.handleColorChange} 
                      style={{ width: '60px', height: '40px' }} 
                    />
                  </StackItem>
                  <StackItem grow>
                    <TextField 
                      value={color}
                      onChange={(e, newValue) => newValue && this.setState({ color: newValue })}
                      placeholder="#0078d4"
                    />
                  </StackItem>
                </Stack>
              </StackItem>
              
              <StackItem>
                <Label>Font Size: {fontSize}px</Label>
                <Slider
                  min={12}
                  max={24}
                  step={1}
                  value={fontSize}
                  onChange={this.handleFontSizeChange}
                  showValue={false}
                />
              </StackItem>

              <StackItem>
                <Label>Site Logo:</Label>
                <input
                  type="file"
                  accept="image/*"
                  onChange={this.handleLogoChange}
                  style={{ marginBottom: '10px' }}
                />
                {(this.state.logoPreview || this.state.logoUrl) && (
                  <div style={{ marginTop: '10px' }}>
                    <Label>Logo Preview:</Label>
                    <img
                      src={this.state.logoPreview || this.state.logoUrl}
                      alt="Logo preview"
                      style={{ 
                        maxWidth: '200px', 
                        maxHeight: '100px', 
                        objectFit: 'contain',
                        border: '1px solid #ccc',
                        borderRadius: '4px',
                        padding: '5px'
                      }}
                    />
                  </div>
                )}
              </StackItem>

              <StackItem>
                <Label>Preview Text:</Label>
                <TextField
                  value={this.state.previewText}
                  onChange={(e, newValue) => newValue !== undefined && this.setState({ previewText: newValue })}
                  placeholder="Enter text to preview"
                />
              </StackItem>
              
              <StackItem>
                <Label>Preview:</Label>
                <div className={combinedPreviewClass}>
                  {this.state.previewText || "Settings Preview"}
                </div>
              </StackItem>
            </Stack>

            <DialogFooter>
              <PrimaryButton 
                onClick={this.handleSaveClick} 
                text={isSaving ? "Saving..." : "Save Changes"}
                disabled={isSaving}
              />
            </DialogFooter>
          </>
        )}
      </Dialog>
    );
  }
}

/**
 * SettingsDialog static class to show the dialog
 */
export class SettingsDialog {
  private static domElement: HTMLDivElement | null = null;
  private static context: ApplicationCustomizerContext | undefined;

  /**
   * Shows the settings dialog
   */
  public static show(context?: ApplicationCustomizerContext): void {
    // Save context for SharePoint operations
    this.context = context;
    
    // Create container div if it doesn't exist
    if (!this.domElement) {
      this.domElement = document.createElement('div');
      document.body.appendChild(this.domElement);
    }

    // Function to dismiss the dialog
    const onDismiss = (): void => {
      if (this.domElement) {
        ReactDOM.unmountComponentAtNode(this.domElement);
      }
    };

    // Render the dialog
    if (this.domElement) {
      ReactDOM.render(
        <SettingsDialogContent onDismiss={onDismiss} context={this.context} />,
        this.domElement
      );
    }
  }
}
