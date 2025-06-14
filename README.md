# Monarch 360 NavBar CRUD - SPFx Application Customizer

## Summary

This SharePoint Framework (SPFx) application customizer adds a settings gear icon to the left of the site logo in SharePoint. When clicked, it opens a dialog that allows users to change the background color and font size of the SharePoint "ShyHeader" element. Settings are stored in and retrieved from a SharePoint list named `navbarcrud`.

## Features

- **Visual Settings Control**: Intuitive UI for customizing SharePoint header appearance
- **Real-time Preview**: Changes apply immediately to the ShyHeader element
- **SharePoint List Integration**: Uses PnPjs to store settings in SharePoint instead of localStorage
- **Modern UI**: Built with Fluent UI React components for consistent SharePoint look and feel
- **Error Handling**: Graceful error handling with user notifications

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

### SharePoint List Setup
Before deploying the solution, you need to create a SharePoint list named `navbarcrud` with the following structure:

- **List Name**: `navbarcrud`
- **Columns**: 
  - `Title` (Single line of text) - Default column
  - `value` (Single line of text) - Custom column to add

**Required List Items**:
- Item 1: Title = "background_color", value = "#0078d4"
- Item 2: Title = "font_size", value = "16"

See [DEPLOYMENT.md](./DEPLOYMENT.md) for detailed setup instructions.

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| spfx-extension | Monarch 360 Development Team |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | June 5, 2025   | Initial release with SharePoint list integration |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
