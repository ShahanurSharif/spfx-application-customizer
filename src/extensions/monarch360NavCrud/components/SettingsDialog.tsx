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
    this.loadSettings();
  }
  
  /**
   * Load settings from localStorage (in a real implementation, this would use SPFx PnPjs or Graph API)
   */
  private loadSettings = (): void => {
    setTimeout(() => {
      try {
        const savedSettings = localStorage.getItem('monarch360Settings');
        
        if (savedSettings) {
          const settings = JSON.parse(savedSettings) as IUserSettings;
          this.setState({
            color: settings.backgroundColor || '#0078d4',
            fontSize: settings.fontSize || 16,
            isLoading: false,
            notification: {
              message: 'Settings loaded successfully',
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
        } else {
          this.setState({ isLoading: false });
        }
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
    }, 1000); // Simulate loading delay
  }
  
  /**
   * Save settings to localStorage (in a real implementation, this would use SPFx PnPjs or Graph API)
   */
  private saveSettings = (settings: IUserSettings): Promise<void> => {
    return new Promise((resolve, reject) => {
      setTimeout(() => {
        try {
          localStorage.setItem('monarch360Settings', JSON.stringify(settings));
          resolve();
        } catch (error) {
          console.error('Error saving settings:', error);
          reject(error);
        }
      }, 1000); // Simulate save delay
    });
  }

  /**
   * Handles dialog close
   */
  private closeDialog = (): void => {
    this.setState({ hideDialog: true });
    this.props.onDismiss();
  }

  /**
   * Handles previous button click
   */
  private handlePrevClick = (): void => {
    console.log('Previous button clicked');
    // Add your previous step logic here
    alert('Going back to previous step');
  }

  /**
   * Applies the settings to the page
   * This is a demonstration function that would be expanded in a real implementation
   */
  private applySettings = (color: string, fontSize: number): void => {
    // In a real implementation, you would:
    // 1. Use SPFx APIs to store settings in user/site properties
    // 2. Apply the styles to the appropriate elements
    
    try {
      // Just for demonstration - apply color to the Suite Bar
      const suiteBar = document.querySelector('.ms-CommandBar');
      if (suiteBar) {
        suiteBar.setAttribute('style', `background-color: ${color} !important`);
      }
      
      // Apply font size to body or specific elements as needed
      document.body.style.setProperty('--default-font-size', `${fontSize}px`);
      
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
      
      // Save settings
      const settings: IUserSettings = {
        backgroundColor: color,
        fontSize: fontSize
      };
      
      await this.saveSettings(settings);
      
      // Apply the settings
      this.applySettings(color, fontSize);
      
      // Show success notification
      this.showNotification('Settings saved successfully!', MessageBarType.success);
      
      // Clear saving state
      this.setState({ isSaving: false });
      
      // Close dialog after a short delay
      setTimeout(() => {
        this.closeDialog();
      }, 1500);
    } catch (error) {
      console.error('Error saving settings:', error);
      this.setState({ isSaving: false });
      this.showNotification('Error saving settings', MessageBarType.error);
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
      title: 'Site Settings',
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
                <Label>Background Color:</Label>
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
                text="Previous" 
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

  /**
   * Shows the settings dialog
   */
  public static show(): void {
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
        <SettingsDialogContent onDismiss={onDismiss} />,
        this.domElement
      );
    }
  }
}
