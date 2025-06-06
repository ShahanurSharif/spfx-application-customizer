# SPFx Extension Deployment Quick Guide

## 🚀 You're Ready to Deploy!

Your SPFx extension with logo functionality is ready for deployment. Here's what you need to do:

### ✅ **Current Status:**
- ✅ Extension built and packaged (`sharepoint/solution/spfx-extension.sppkg`)
- ✅ Logo functionality implemented in settings dialog
- ✅ Playwright E2E tests configured and passing
- ✅ Dynamic configuration loading from `config/serve.json`

### 🔧 **Pre-Deployment: Create SharePoint List**

First, create the required SharePoint list by running the PowerShell script:

```powershell
# Run this in PowerShell with admin rights
.\setup-sharepoint-list.ps1 -SiteUrl "https://monarch360demo.sharepoint.com/sites/shan"
```

**Or manually create the list:**
1. Go to SharePoint site: `https://monarch360demo.sharepoint.com/sites/shan`
2. Create a new list named: `navbarcrud`
3. Add a column: `value` (Single line of text)
4. Add these items:
   - Title: `background_color`, value: `#0078d4`
   - Title: `font_size`, value: `16`
   - Title: `logo`, value: `` (empty for now)

### 📦 **Deployment Steps:**

#### Option 1: SharePoint Admin Center (Recommended)
1. Go to SharePoint Admin Center
2. Navigate to "Apps" → "App Catalog"
3. Upload `sharepoint/solution/spfx-extension.sppkg`
4. Deploy the solution
5. Go to your site and activate the extension

#### Option 2: Site Collection App Catalog
1. Go to your SharePoint site: `https://monarch360demo.sharepoint.com/sites/shan`
2. Site Settings → Site Collection Features
3. Activate "Site Collection App Catalog" if not already active
4. Upload the `.sppkg` file to the App Catalog
5. Deploy and install the extension

### 🧪 **Testing After Deployment:**

Once deployed, run the tests again to see full functionality:

```bash
# Run all E2E tests
npx playwright test

# Run with UI for interactive testing
npx playwright test --ui

# Run specific logo tests
npx playwright test --grep "logo"
```

### 🎨 **Using the Extension:**

After deployment, you'll see:
1. ⚙️ **Settings gear icon** next to the SharePoint site logo
2. **Settings dialog** with:
   - Background color picker
   - Font size slider (12-24px)
   - Logo URL input field
3. **Live preview** of changes
4. **Save to SharePoint list** functionality

### 🔧 **Logo Functionality:**
- Enter any logo URL in the settings dialog
- Supports: PNG, JPG, SVG, GIF formats
- Logo will replace the SharePoint site logo
- Settings are stored in the `navbarcrud` SharePoint list

### 📊 **SharePoint List Structure:**
```
navbarcrud list:
├── Title: "background_color" → value: "#0078d4"
├── Title: "font_size" → value: "16"
└── Title: "logo" → value: "https://example.com/logo.png"
```

---

**🎯 Next Steps:**
1. Run the PowerShell script to create the SharePoint list
2. Deploy the `.sppkg` file to SharePoint
3. Test the extension functionality
4. Customize settings through the dialog

**Need help?** The tests will guide you - they show exactly what to expect at each step!
