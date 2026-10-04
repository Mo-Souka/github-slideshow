# Goods Receiving web part: deployment guide for IT

This guide is for the IT / Microsoft 365 administrators who install the **Goods Receiving** web
part for the warehouse. No programming knowledge is needed.

**Estimated effort:** 30–60 minutes, once.

---

## 1. What is being installed (summary for approval)

| Item | Details |
|---|---|
| Type | SharePoint Framework (SPFx) web part. Client-side code that runs inside SharePoint Online pages. |
| File | `goods-receiving.sppkg` (built with `npm run build`, found in `sharepoint/solution/`) |
| Framework version | SPFx 1.23.2 (Microsoft's current stable release) |
| Where data is stored | Only in SharePoint Online, on the warehouse site: the list **Material Receiving Records** and the library **Receiving Documents** |
| Sign-in | The users' existing Microsoft 365 accounts. No separate login. |
| External servers / Azure | **None.** The code is packaged inside the .sppkg and served by SharePoint (from the Microsoft 365 CDN if that is enabled in your tenant). |
| API permissions (Microsoft Graph etc.) | **None requested.** The web part only calls the SharePoint REST API of the site it runs on, as the signed-in user. Users can never see or change more than their SharePoint permissions allow. |
| "Custom script" required | No |
| Licences | Nothing beyond SharePoint Online |

---

## 2. What is the App Catalog?

The **App Catalog** is a special SharePoint site where custom apps and web parts (`.sppkg` files)
are uploaded and approved before anyone can use them. There are two kinds:

| | Tenant App Catalog | Site Collection App Catalog |
|---|---|---|
| Scope | Whole organisation (one per tenant) | Only one site (e.g. the warehouse site) |
| Who sets it up | SharePoint Administrator | SharePoint Administrator enables it for a site |
| Who uploads apps | SharePoint Administrator, or an owner of the App Catalog site | Site collection administrators of that site |
| When to use | Standard choice | If you want the web part available on the warehouse site only, or the tenant catalog is tightly controlled |

**Does your tenant already have a tenant App Catalog?**

1. Open the **SharePoint admin center** (`https://<tenant>-admin.sharepoint.com`).
2. Go to **More features → Apps → Open**.
3. If you see the **Manage apps** page, the App Catalog exists. In current tenants it is created
   automatically the first time an administrator opens this page. Older tenants often have it at
   `https://<tenant>.sharepoint.com/sites/appcatalog`.

---

## 3. Permissions required

| Task | Who can do it |
|---|---|
| Open SharePoint admin center / create the tenant App Catalog | **SharePoint Administrator** (or Global Administrator) |
| Upload and enable the app in the **tenant** App Catalog | SharePoint Administrator, or an **owner of the App Catalog site** |
| Enable a **site collection** App Catalog | SharePoint Administrator |
| Upload the app to a site collection App Catalog | **Site collection administrator** of that site |
| Register the Entra ID app for PnP PowerShell (one time, section 5) | **Application Administrator**, Cloud Application Administrator or Global Administrator (to grant consent) |
| Run the provisioning script | **Site owner** (Full Control) of the warehouse site |
| Add people to the two SharePoint groups | Site owner |
| Add the web part to a page | Anyone who can **edit pages** on the site (site members by default) |

---

## 4. Deploy the package

### Option A: Tenant App Catalog (recommended)

1. SharePoint admin center → **More features → Apps → Open**.
2. Click **Upload** and select `goods-receiving.sppkg`.
3. A panel **"Enable app"** appears. Choose **"Only enable this app"**, then click **Enable app**.
   - The web part then becomes *available* to add on any site. It is not added to any page automatically.
   - No "API access" request appears (none is needed).
4. Done. It can take a few minutes until the web part shows up in the page toolbox.

PowerShell alternative (PnP PowerShell, run as SharePoint Administrator):

```powershell
Connect-PnPOnline -Url https://<tenant>.sharepoint.com/sites/appcatalog -Interactive -ClientId <app id>
Add-PnPApp -Path .\goods-receiving.sppkg -Scope Tenant -Publish -SkipFeatureDeployment -Overwrite
```

### Option B: Site Collection App Catalog (only the warehouse site)

1. A SharePoint Administrator enables it once:
   ```powershell
   Connect-PnPOnline -Url https://<tenant>-admin.sharepoint.com -Interactive -ClientId <app id>
   Add-PnPSiteCollectionAppCatalog -Site https://<tenant>.sharepoint.com/sites/warehouse
   ```
   (or with the SharePoint Online Management Shell: `Add-SPOSiteCollectionAppCatalog -Site <url>`)
2. On the warehouse site go to **Site contents → Apps for SharePoint**.
3. Upload `goods-receiving.sppkg`, and click **Deploy** / **Enable app** when asked.
4. If the web part does not appear in the page toolbox, add the app to the site:
   **Site contents → + New → App → Goods Receiving**.

---

## 5. Provision the list, library and groups (once per site)

The script `provisioning/Provision-GoodsReceiving.ps1` creates everything the web part needs on
the warehouse site:

- the list and library
- all columns and indexes
- version history settings
- the groups *Receiving Team* and *Receiving Supervisors*
- the permissions for those groups
- the **Goods Receiving page**, with the web part on it and a navigation link

It is safe to run again; anything that exists is left unchanged.

> **Order matters for the page:** deploy the package first (section 4), then run the script. If the
> script runs before the package is deployed, it creates everything except the page, and tells
> you to run it again afterwards.

### 5.1 Prerequisites on the computer that runs the script

- **PowerShell 7.4 or newer**: `winget install Microsoft.PowerShell`
- **PnP.PowerShell**: `Install-Module PnP.PowerShell -Scope CurrentUser`

### 5.2 One-time: an Entra ID app registration for PnP PowerShell

Since September 2024, PnP PowerShell no longer provides a shared sign-in app. Each organisation
registers its own. An Application Administrator or Global Administrator runs:

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell" -Tenant <tenant>.onmicrosoft.com -Interactive
```

- This creates an app registration with **delegated** permissions only.
- The script can never do more than the person signed in to it is allowed to do.
- Note the **Application (client) ID** it prints. You need it below.
- If your organisation already has a PnP PowerShell app registration, use that one.

### 5.3 Run the script

From the `goods-receiving` folder, signed in as a **site owner** of the warehouse site:

```powershell
./provisioning/Provision-GoodsReceiving.ps1 `
    -SiteUrl https://<tenant>.sharepoint.com/sites/warehouse `
    -ClientId <Application (client) ID from 5.2>
```

The script prints what it created (`[created]`), what already existed (`[exists]`) and what it
changed (`[updated]`).

**What it changes on permissions:** the list and the library stop inheriting permissions from the
site. All existing access is copied, so site owners, members and visitors keep the access they
had. Then:

- *Receiving Team* gets **Contribute**.
- *Receiving Supervisors* get a new permission level, **Receiving Supervisor**: Contribute +
  Approve Items + Override List Behaviors.

If you prefer to manage permissions yourself, add `-SkipPermissions`.

**Recommended:** to make sure only the two groups can change records, remove (or set to Read) the
site's *Members* group on the list and library afterwards. See README.md, *Security model*, for
further hardening options.

### 5.4 After the script

1. **Add people to the groups:** Site settings → People and groups (or the site's *Settings → Site permissions → Advanced permissions settings*):
   - **Receiving Team:** warehouse staff who record shipments.
   - **Receiving Supervisors:** people who approve or reject records.
   - You can add people directly or add a Microsoft 365 / security group.
2. **Check the site's time zone:** Site settings → **Regional settings → Time zone** must be the
   plant's local time zone. Otherwise receiving dates can appear one day off.

---

## 6. The Goods Receiving page

The provisioning script creates and publishes the page **Goods Receiving**
(`https://<tenant>.sharepoint.com/sites/warehouse/SitePages/Goods-Receiving.aspx`):

- a banner with the page title
- a short intro text
- the web part below it
- a link at the top of the site navigation

Nothing else is needed. Share the link with the warehouse team. On tablets and phones the page
works in the browser and in the SharePoint mobile app.

To change the page title or intro text before running the script, edit the `page` section in
`src/config/solution.json`. After the page exists, edit it in SharePoint like any other page.
The script never changes an existing page.

### Adding the web part to another page manually

1. On the warehouse site click **+ New → Page**. Choose a blank page and name it, e.g. **Goods Receiving**.
2. Use a **one-column, full-width or wide section**, which gives the most room for the table.
3. Click **+** in the section, search for **Goods Receiving**, and add it.
4. Optional: click the pencil icon on the web part to open its settings:
   - **Records list / Documents library:** leave empty (defaults are correct).
   - **Site URL:** only if the list lives on another site.
   - **Default date range** (days) and **Records per page**.
5. Click **Publish**.
6. Optional: add the page to the site navigation and share the link with the warehouse team.
   On tablets and phones the page works in the browser and in the SharePoint mobile app.

**Full-page app (optional):** the web part also supports full-page mode. This hides the page
title area and gives the app the whole screen, which is useful on warehouse tablets. A site owner
can create such a page with PnP PowerShell:

```powershell
Add-PnPPage -Name "GoodsReceivingApp" -LayoutType SingleWebPartAppPage
Add-PnPPageWebPart -Page "GoodsReceivingApp" -Component "Goods Receiving"
Set-PnPPage -Identity "GoodsReceivingApp" -Publish
```

---

## 7. Updating to a new version

1. The developer increases the version in `config/package-solution.json` and builds a new `goods-receiving.sppkg`.
2. Upload it to the same App Catalog. When asked, choose **Replace**, then **Enable app** again.
3. Existing pages use the new version automatically. Browsers may need a refresh, and it can take a few minutes.
4. If the new version added fields, a site owner re-runs the provisioning script (section 5.3).

## 8. Removing the app

- Delete `goods-receiving.sppkg` from the App Catalog. Pages then show an empty placeholder where the web part was.
- **The data is not deleted.** The list, library, versions and groups stay, and can be removed
  separately by a site owner if wanted.

---

## 9. Checklist

- [ ] App Catalog exists (tenant, or site collection for the warehouse site)
- [ ] `goods-receiving.sppkg` uploaded and enabled ("Only enable this app")
- [ ] Entra ID app registration for PnP PowerShell available (client ID known)
- [ ] Provisioning script run on the warehouse site without errors
- [ ] People added to *Receiving Team* and *Receiving Supervisors*
- [ ] Site time zone = plant time zone
- [ ] Provisioning script output says the page **Goods Receiving** was created (or already exists)
- [ ] Smoke test:
  - [ ] a receiver creates and submits a record with a photo
  - [ ] a different supervisor approves it
  - [ ] Export to Excel works
