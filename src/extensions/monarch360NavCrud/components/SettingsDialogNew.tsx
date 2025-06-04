import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { 
  Dialog, 
  DialogType, 
  IDialogContentProps, 
  DialogFooter,
  PrimaryButton, 
  DefaultButton, 
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

// Context for SharePoint operations
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';

/**
 * Settings Interface
 */
export interface IUserSettings {
  backgroundColor: string;
  fontSize: number;
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
        this.setState({ isLoading: false });
        return;
      }
      
      // Initialize SP
      const sp = spfi().using(SPFx(this.props.context));
      
      // Get background color from list
      const bgColorItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'background_color'")
        .top(1)();
        
      // Get font size from list
      const fontSizeItems = await sp.web.lists.getByTitle("navbarcrud").items
        .filter("Title eq 'font_size'")
        .top(1)();
        
      // Default values
      let backgroundColor = '#0078d4';
      let fontSize = 16;
      
      // If items exist, use those values
      if (bgColorItems.length > 0) {
        backgroundColor = bgColorItems[0].value || backgroundColor;
      }
      
      if (fontSizeItems.length > 0) {
        fontSize = parseInt(fontSizeItems[0].value, 10) || fontSize;
      }
      
      this.setState({
        color: backgroundColor,
        fontSize: fontSize,
        isLoading: false,
        notification: {
          message: 'Settings loaded successfully from SharePoint list',
          type: MessageBarType.success,
          show: true
        }
      });
      
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
   * Handles preview button click
   */
  private handlePrevClick = (): void => {
    const { color, fontSize } = this.state;
    this.applySettings(color, fontSize);
    this.showNotification('Preview applied! Click Save to keep these settings.', MessageBarType.info);
  }

  /**
   * Applies the settings to the page
   */
  private applySettings = (color: string, fontSize: number): void => {
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
      
      console.log(`Applied settings - Color: ${color}, Font Size: ${fontSize}px`);
    } catch (error) {
      console.error('Failed to apply settings:', error);
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
    const { color, fontSize } = this.state;
    
    try {
      // Set saving state
      this.setState({ isSaving: true });
      
      // Save settings to SharePoint list
      const settings: IUserSettings = {
        backgroundColor: color,
        fontSize: fontSize
      };
      
      await this.saveSettings(settings);
      
      // Apply the settings
      this.applySettings(color, fontSize);
      
      // Show success notification
      this.showNotification('Settings saved successfully to SharePoint list!', MessageBarType.success);
      
      // Clear saving state
      this.setState({ isSaving: false });
      
      // Close dialog after a short delay
      setTimeout(() => {
        this.closeDialog();
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
              <DefaultButton 
                onClick={this.handlePrevClick}
                text="Preview" 
                disabled={isSaving}
              />
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
