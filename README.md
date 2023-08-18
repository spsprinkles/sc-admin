[![Build & Package](https://github.com/spsprinkles/sc-admin/actions/workflows/webpack.yml/badge.svg)](https://github.com/spsprinkles/sc-admin/actions/workflows/webpack.yml)

# Site Administration Tool
### A tool to help manage SharePoint sites.
![SiteAdminTool](https://github.com/spsprinkles/sc-admin/assets/24440567/802d2a12-bf90-4ed6-a9d3-68b80690bb9d)
![SiteAdminTool2](https://github.com/spsprinkles/sc-admin/assets/24440567/c89d4a05-8042-47da-8f35-94ecaf0ee052)

### Download the latest [release](https://github.com/spsprinkles/sc-admin/releases/tag/v0.1.0) as an SPFx webpart and deploy it directly to SharePoint Online.

## Building the Solution (for on-premises SharePoint)

### Required Software

* NodeJS 16.x (LTS)
* Code Editor (Visual Studio Code)

### Install Libraries

Run `npm i` to install the required libraries.

### Build the Code

Run `npm run build` to build the dev version. Run `npm run prod` to build the minified version.

## Install the Solution

### Copy Assets

Copy the `dist/sc-admin.min.js` file to the `Site Assets` library.

### Create List

#### Requirements

The solution is designed to work in classic pages. Click the link in the bottom left to view modern pages in classic mode.

#### Access the Developer Tools

Press `Ctrl+Shift+I` or `F-12` to access the developer tools and access the console tab.

#### Reference the Solution

Type in the following to reference the script.

```js
var el = document.createElement("script");
el.src = "/sites/dev/siteassets/sc-admin.min.js";
document.head.appendChild(el);
```
_Update the src url w/ your target web._

#### Install the Solution

```js
SCAdmin.Configuration.install();
```

Validate the logs from the installation to validate the list was created.

#### Access the List

View the site contents and access the `Site Collections` list. Setup the list to default to `Classic` mode. The administration menu will be displayed on the default view.
