# PowerShell script to create the required SharePoint list for Monarch360 NavBar CRUD extension
# Run this script in PowerShell with PnP PowerShell installed

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$ListName = "navbarcrud"
)

# Check if PnP PowerShell is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "❌ PnP PowerShell module is not installed." -ForegroundColor Red
    Write-Host "📦 Please install it first by running:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name PnP.PowerShell -Scope CurrentUser" -ForegroundColor Cyan
    exit 1
}

try {
    Write-Host "🔗 Connecting to SharePoint site: $SiteUrl" -ForegroundColor Green
    
    # Connect to SharePoint
    Connect-PnPOnline -Url $SiteUrl -Interactive
    
    Write-Host "✅ Connected successfully!" -ForegroundColor Green
    
    # Check if list already exists
    $existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($existingList) {
        Write-Host "⚠️ List '$ListName' already exists. Checking structure..." -ForegroundColor Yellow
        
        # Check if the 'value' column exists
        $valueField = Get-PnPField -List $ListName -Identity "value" -ErrorAction SilentlyContinue
        if (-not $valueField) {
            Write-Host "📝 Adding missing 'value' column..." -ForegroundColor Blue
            Add-PnPField -List $ListName -DisplayName "value" -InternalName "value" -Type Text
            Write-Host "✅ 'value' column added successfully!" -ForegroundColor Green
        } else {
            Write-Host "✅ 'value' column already exists." -ForegroundColor Green
        }
    } else {
        Write-Host "📋 Creating list '$ListName'..." -ForegroundColor Blue
        
        # Create the list
        New-PnPList -Title $ListName -Template GenericList -EnableContentTypes:$false
        Write-Host "✅ List created successfully!" -ForegroundColor Green
        
        # Add the 'value' column
        Write-Host "📝 Adding 'value' column..." -ForegroundColor Blue
        Add-PnPField -List $ListName -DisplayName "value" -InternalName "value" -Type Text
        Write-Host "✅ 'value' column added successfully!" -ForegroundColor Green
    }
    
    # Check existing items
    $existingItems = Get-PnPListItem -List $ListName
    $backgroundColorItem = $existingItems | Where-Object { $_.FieldValues.Title -eq "background_color" }
    $fontSizeItem = $existingItems | Where-Object { $_.FieldValues.Title -eq "font_size" }
    $logoItem = $existingItems | Where-Object { $_.FieldValues.Title -eq "logo" }
    
    # Add background_color item if it doesn't exist
    if (-not $backgroundColorItem) {
        Write-Host "🎨 Adding 'background_color' item..." -ForegroundColor Blue
        Add-PnPListItem -List $ListName -Values @{
            "Title" = "background_color"
            "value" = "#0078d4"
        }
        Write-Host "✅ Background color item added (default: #0078d4)" -ForegroundColor Green
    } else {
        Write-Host "✅ Background color item already exists: $($backgroundColorItem.FieldValues.value)" -ForegroundColor Green
    }
    
    # Add font_size item if it doesn't exist
    if (-not $fontSizeItem) {
        Write-Host "📏 Adding 'font_size' item..." -ForegroundColor Blue
        Add-PnPListItem -List $ListName -Values @{
            "Title" = "font_size"
            "value" = "16"
        }
        Write-Host "✅ Font size item added (default: 16px)" -ForegroundColor Green
    } else {
        Write-Host "✅ Font size item already exists: $($fontSizeItem.FieldValues.value)px" -ForegroundColor Green
    }
    
    # Add logo item if it doesn't exist
    if (-not $logoItem) {
        Write-Host "🖼️ Adding 'logo' item..." -ForegroundColor Blue
        Add-PnPListItem -List $ListName -Values @{
            "Title" = "logo"
            "value" = ""
        }
        Write-Host "✅ Logo item added (ready for image upload)" -ForegroundColor Green
    } else {
        Write-Host "✅ Logo item already exists" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "🎉 Setup completed successfully!" -ForegroundColor Green
    Write-Host "📋 List '$ListName' is ready with the following items:" -ForegroundColor Yellow
    
    # Display final list contents
    $finalItems = Get-PnPListItem -List $ListName
    foreach ($item in $finalItems) {
        Write-Host "   - $($item.FieldValues.Title): $($item.FieldValues.value)" -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "🚀 You can now test your SPFx extension!" -ForegroundColor Green
    
} catch {
    Write-Host "❌ Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "💡 Make sure you have the necessary permissions on the SharePoint site." -ForegroundColor Yellow
} finally {
    # Disconnect
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "🔌 Disconnected from SharePoint." -ForegroundColor Gray
    } catch {
        # Ignore disconnect errors
    }
}

Write-Host ""
Write-Host "📖 Usage examples:" -ForegroundColor Yellow
Write-Host "   .\setup-sharepoint-list.ps1 -SiteUrl 'https://yourtenant.sharepoint.com/sites/yoursite'" -ForegroundColor Cyan
Write-Host "   .\setup-sharepoint-list.ps1 -SiteUrl 'https://yourtenant.sharepoint.com/sites/yoursite' -ListName 'customlistname'" -ForegroundColor Cyan
