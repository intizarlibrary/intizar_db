# Intizarul Imamul Muntazar Database

A static GitHub Pages frontend for managing members, graduates, and Mas'ulin. The frontend communicates with a Google Apps Script web app, which stores records in Google Sheets.

## Project Structure

- `index.html` - Public landing page and login dialog.
- `dashboard.html` - Dashboard, registries, reports, filters, and administration.
- `registration.html` - Member and Mas'ul registration forms.
- `script.js` - Frontend state, API requests, rendering, authentication, and actions.
- `style.css` - Shared application styling.
- `Code.gs` - Google Apps Script request dispatcher.
- `Utils.gs` - Google Sheets business logic, permissions, validation, and audit logging.

## Recent Changes

### Record deletion

- Admin users can delete member records from the Members registry.
- Admin users can delete Mas'ul records from the Mas'ulin registry.
- Every delete requires confirmation in the dashboard.
- The backend enforces Admin-only access; hiding the button is not the security control.
- Successful deletions are written to the AuditLog sheet as `MEMBER_DELETED` or `MASUL_DELETED`.
- The relevant table, graduate list, and dashboard statistics refresh after deletion.
- If the last record on a page is deleted, the list automatically moves back to the previous page.

Deletion is permanent from the relevant Google Sheet. The audit entry keeps the deleted ID and name, but the full deleted row is not archived.

### Reliability fixes

- Mas'ul search now safely normalizes numbers, blanks, and strings returned by Google Sheets.
- Added a visible delete icon and destructive-action styling to the dashboard.
- Added `deleteMember` and `deleteMasul` actions to the Apps Script dispatcher.

## Deployment

### GitHub Pages frontend

1. Push the HTML, CSS, JavaScript, image, and metadata files to the GitHub repository.
2. Enable GitHub Pages for the branch and folder containing these files.
3. Open the published Pages URL and test login, dashboard loading, registration, editing, and searching.

The Apps Script URL is configured near the top of `script.js` in `APPS_SCRIPT_URL`.

### Google Apps Script backend

After changing `Code.gs` or `Utils.gs`:

1. Open the Apps Script project connected to the database.
2. Copy or sync the updated `Code.gs` and `Utils.gs` code.
3. Save the project.
4. Deploy a new version of the web app.
5. Set access to the required users, normally `Anyone` for a public GitHub Pages frontend.
6. Keep the deployment URL the same, or update `APPS_SCRIPT_URL` in `script.js` if it changes.

The frontend cannot activate backend changes by itself. The new delete actions will return an unknown-action error until the updated Apps Script version is deployed.

## Backend Smoke Test

The configured backend supports a read-only ping request:

```text
https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?payload=%7B%22action%22%3A%22ping%22%7D
```

A healthy response includes:

```json
{"success":true,"message":"System online"}
```

## Permissions

- `Admin` can manage records, Mas'ulin, zones, branches, settings, exports, and audit logs.
- `Zonal Mas'ul` can work with records within the assigned zone.
- `Branch Mas'ul` can register and view records within the assigned branch.
- Record deletion is restricted to `Admin` in the backend.

## Validation

The current changes were checked with:

- JavaScript syntax validation for `script.js`.
- JavaScript parsing validation for `Code.gs` and `Utils.gs`.
- Workspace diagnostics with no reported errors.
- A live read-only `ping` request to the configured Apps Script endpoint.
