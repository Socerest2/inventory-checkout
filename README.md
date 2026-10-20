# Inventory Check Out

A check-out/check-in page for the **26-27 Warehouse Inventory** Google Sheet. The page is hosted on GitHub Pages, and a Google Apps Script handles the sheet.

- Every check-out adds a row to a **Checkout Log** tab in the sheet. A check-in marks that row `IN`.
- Items are read live from each office's sheet tab (columns A–C).

## Admin portal

The **Admin** link at the bottom of the site opens the admin portal. You can change everything there without editing any code:

- **Checkouts:** search the full log, check items in, and edit or delete entries.
- **Inventory:** edit, add or delete item rows in each office's sheet tab.
- **Settings:** set the site title, choose whether an event is required, and manage offices (add, rename, reorder, or point them at a different sheet tab and starting row).

**Who can open the portal:** the people on the admin list in **Settings → Admins**. Admins add and remove each other there. If someone you add doesn't have edit access to the sheet yet, adding them gives it to them, because the portal needs it. If the list is ever empty, every editor of the sheet can get in.

## Apps Script setup

The Apps Script project has two files, [`apps-script/Code.gs`](apps-script/Code.gs) and [`apps-script/Admin.html`](apps-script/Admin.html). It is deployed twice:

| Deployment | Execute as | Who has access | Used by |
|---|---|---|---|
| Public API | Me | Anyone | `API_URL` in `config.js` |
| Admin portal | User accessing the web app | Anyone with Google account | `ADMIN_URL` in `config.js` (with `?page=admin`) |

After you change the script, update **both** deployments with **Deploy → Manage deployments → Edit → Version: New version**. Their URLs stay the same.
