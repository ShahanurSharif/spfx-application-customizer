# SPFx Application Customizer - Deployment Guide

## Overview
This SharePoint Framework (SPFx) application customizer adds a settings gear icon to the left of the site logo in SharePoint. When clicked, it opens a dialog to change the background color and font size of the SharePoint "ShyHeader" element.

## Prerequisites

### SharePoint List Setup
Before deploying the solution, you need to create a SharePoint list named `navbarcrud` with the following configuration:

1. **List Name**: `navbarcrud`
2. **Columns**:
   - `Title` (Single line of text) - Default column, already exists
   - `value` (Single line of text) - Add this column

3. **Required List Items**:
   Add these two items to the list:
   - Item 1: Title = "background_color", value = "#0078d4" (or any valid CSS color)
   - Item 2: Title = "font_size", value = "16" (or any valid font size in pixels)

### PowerShell Commands to Create the List
```powershell
# Connect to your SharePoint site
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive

# Create the list
New-PnPList -Title "navbarcrud" -Template GenericList

# Add the value column
Add-PnPField -List "navbarcrud" -DisplayName "value" -InternalName "value" -Type Text

# Add default items
Add-PnPListItem -List "navbarcrud" -Values @{"Title"="background_color";"value"="#0078d4"}
Add-PnPListItem -List "navbarcrud" -Values @{"Title"="font_size";"value"="16"}
```

## Deployment Steps

### 1. Upload the Solution Package
1. Navigate to your SharePoint Admin Center or Site Collection App Catalog
2. Upload the `spfx-extension.sppkg` file located in the `sharepoint/solution/` folder
3. Deploy the solution when prompted

### 2. Install the Extension on Your Site
1. Go to your SharePoint site
2. Navigate to Site Settings > Site Collection Features (or Site Features)
3. Activate the "Application Extension - Deployment of custom action" feature

### Alternative: Direct Installation via PowerShell
```powershell
# Connect to your SharePoint site
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive

# Install the app
Install-PnPApp -Identity "spfx-extension" -Scope Site
```

## Features

### Settings Gear Icon
- **Location**: Positioned to the left of the SharePoint site logo
- **Style**: Uses Fluent UI icons with hover effects
- **Functionality**: Opens a settings dialog when clicked

### Settings Dialog
- **Background Color**: Choose from predefined colors or enter a custom hex color
- **Font Size**: Adjust the font size using a slider (10-24px range)
- **Real-time Preview**: Changes are applied immediately to the ShyHeader
- **Persistence**: Settings are saved to and retrieved from the SharePoint `navbarcrud` list

### SharePoint Integration
- **Data Storage**: Uses SharePoint list instead of browser localStorage
- **PnPjs Integration**: Utilizes @pnp/sp for efficient SharePoint operations
- **Error Handling**: Graceful error handling with user notifications

## Technical Details

### Key Files
- `Monarch360NavCrudApplicationCustomizer.ts` - Main application customizer
- `SettingsDialogNew.tsx` - React component for the settings dialog
- `spfx-extension.sppkg` - Deployable solution package

### Dependencies
- **@pnp/sp**: SharePoint REST API integration
- **@pnp/logging**: Logging utilities
- **@pnp/common**: Common utilities
- **@fluentui/react**: UI components

### Browser Compatibility
- Modern browsers supporting ES6+
- SharePoint Online
- SharePoint 2019 (with compatibility mode)

## Testing

### Manual Testing Steps
1. **Verify Icon Placement**: Check that the settings gear appears to the left of the site logo
2. **Dialog Functionality**: Click the icon to ensure the dialog opens properly
3. **Settings Loading**: Verify that existing settings load from the SharePoint list
4. **Color Changes**: Test background color changes and verify they apply to ShyHeader
5. **Font Size Changes**: Test font size slider and verify changes apply immediately
6. **Save Functionality**: Confirm that settings are saved to the SharePoint list
7. **Error Handling**: Test with missing list or invalid data to ensure graceful error handling

### Troubleshooting

#### Common Issues
1. **Settings gear not visible**: 
   - Ensure the feature is activated
   - Check browser console for JavaScript errors
   - Verify the extension is properly installed

2. **Dialog not opening**:
   - Check if the `navbarcrud` list exists
   - Verify list permissions
   - Check browser console for errors

3. **Settings not saving**:
   - Verify the `value` column exists in the list
   - Check user permissions to the list
   - Review network requests in browser developer tools

#### Debug Mode
To enable debug mode, add `?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js` to your SharePoint URL when testing locally.

## Support
For issues or questions, please refer to the project repository or contact the development team.
