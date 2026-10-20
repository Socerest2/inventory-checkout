# Inventory Check Out

A check-out/check-in page for the **26-27 Warehouse Inventory** Google Sheet. It's hosted on GitHub Pages and has a Google Apps Script backend.

- Every check-out adds a row to a **Checkout Log** tab in the sheet. The tab is created automatically.
- A check-in marks that row `IN` and records the time.
- Items are read live from the `OFSA Basement Inventory` and `DM Office` tabs.

## One-time setup (about 2 minutes)

1. Go to <https://script.google.com>, click **New project**, and replace `Code.gs` with [`apps-script/Code.gs`](apps-script/Code.gs).
2. Click **Deploy → New deployment → Web app**, and set:
   - **Execute as:** Me
   - **Who has access:** Anyone
3. Click **Deploy**, authorize with your Google account, and copy the **Web app URL**. It ends in `/exec`.
4. Paste that URL into [`config.js`](config.js) as `window.API_URL = '...'`, then commit.

If you edit `Code.gs` later, go to **Deploy → Manage deployments → Edit → New version** so the URL stays the same.

## Autofill

- Typing in **Filter items** narrows the item dropdown. When only one item matches, it's selected automatically. Clicking an item in the list selects it.
- The Name, NetID and Event fields remember what you entered last.
- Typing a NetID that has been used before fills in the matching name. Earlier events are offered as suggestions.

## Adding an office tab

Add an entry to `OFFICES` in `Code.gs` and to `OFFICES` in `index.html`.
