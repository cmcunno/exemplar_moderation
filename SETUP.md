# Moderation Hub for Microsoft Teams: setup guide

About 45 minutes. Do the steps in order.

This guide uses your existing **science_Team** in Teams. Every Team has its own SharePoint site, so the app stores its data there.

**You need:**
- to be an **owner** of science_Team;
- an admin role that can register apps and grant admin consent in Microsoft Entra: Global Administrator, Application Administrator or Cloud Application Administrator;
- permission to upload custom apps to Teams;
- a free GitHub account, to put the app online.

**Where things live:**
- Questions, marks, staff and pasted images are stored in **your school's SharePoint**.
- GitHub only hosts the app's code. It holds no student data and no passwords.

---

## 1. Get your team's SharePoint address and sort out who's an owner

1. In Teams, open **science_Team**, go to any channel and select the **Files** tab.
2. Select **⋯ (or Open) → Open in SharePoint**. A browser opens on the team's SharePoint site.
3. Copy the address **up to and including the site name** and stop there. For example:
   - from `https://yourschool.sharepoint.com/sites/science_Team/Shared%20Documents/...`
   - copy `https://yourschool.sharepoint.com/sites/science_Team`

   You'll need it in step 5.
4. Check who's an owner of the team. In Teams, select **⋯ next to science_Team → Manage team → Members**:
   - **You (the CC) and every Head of Subject** must be **Owner**. Use the **Role** drop-down to change it.
   - **Teachers** stay as **Member**.

   Owners can change everything in the app's lists. Members will only be able to save their own marks once you've done step 7.

## 2. Put the app online (GitHub Pages)

1. Unzip **moderation-hub-files.zip** on your computer.
2. On github.com, select **+ → New repository**.
   - Name: `moderation-hub`
   - Visibility: **Public** (GitHub Pages is only free for public repositories; nothing private is in the code)
   - Select **Create repository**.
3. On the new repository page, select **uploading an existing file**.
4. Drag in **everything inside** the unzipped `moderation-hub-teams` folder: all the files *and* the `lib` folder. Don't drag the outer folder itself. Select **Commit changes**.
5. Select **Settings → Pages**. Under **Branch**, choose **main** and **/ (root)**, then select **Save**.
6. Wait 1–2 minutes and refresh. GitHub shows your address: `https://YOUR-GITHUB-NAME.github.io/moderation-hub/`.
7. Check it works by opening `https://YOUR-GITHUB-NAME.github.io/moderation-hub/package.html`. This is the **setup helper**; keep it open.

## 3. Register the app in Microsoft Entra

1. Go to **entra.microsoft.com → Applications → App registrations → New registration**.
   - Name: `Moderation Hub`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: choose **Single-page application (SPA)** and paste the **first** address from the setup helper (ends in `/index.html`).
   - Select **Register**.
2. On the **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID**. You'll need them in step 5.
3. Select **API permissions → Add a permission → Microsoft Graph → Delegated permissions**. Tick:
   - `Sites.ReadWrite.All`
   - `Sites.Manage.All`

   Select **Add permissions**. `User.Read` is already there; leave it.
4. Select **Grant admin consent for [your school] → Yes**. All three permissions should show a green tick.

> The app acts as whoever is signed in. It can only see or change what that person is already allowed to in SharePoint, which is why the groups in step 1 matter.

## 4. Add the Teams sign-in address

1. In the app registration, select **Authentication**.
2. Under **Single-page application**, select **Add URI** and paste the **second** address from the setup helper. It starts with `brk-multihub://`.
3. Select **Save**.

You should now have two addresses listed under Single-page application.

## 5. Fill in config.js

1. In your GitHub repository, open **config.js** and select the **pencil (Edit)** icon.
2. Replace the four values marked EDIT:
   - `clientId`: the Application (client) ID from step 3
   - `tenantId`: the Directory (tenant) ID from step 3
   - `siteUrl`: the SharePoint site address from step 1
   - `adminEmails`: your school email (the CC). Add more inside the brackets separated by commas, e.g. `["a@school.org.uk", "b@school.org.uk"]`.

   Keep the quote marks. Don't change `appId`.
3. Select **Commit changes**.
4. Wait a minute, then reload the setup helper. Every line under **config.js check** should show ✓.

## 6. Add the app to Teams

1. In the setup helper, select **Download Teams app package**. This saves `moderation-hub-teams.zip`; don't unzip it.
2. Make it available to all staff: go to **Teams admin centre → Teams apps → Manage apps → Upload new app**, and upload the zip.
   - If you can't use the admin centre, use the Teams app instead: **Apps → Manage your apps → Upload an app → Upload a custom app**.
3. Add it to **science_Team**: open a channel (for example General), select **+** (Add a tab) → **Moderation Hub** → **Save**.
   - Staff can also pin it to their left-hand rail from **Apps**.

## 7. First run: create the lists and lock them down

1. Open the Moderation Hub tab yourself and select **Create the SharePoint lists**. This creates five lists in the team's site:
   - Moderation Staff
   - Moderation Questions
   - Moderation Marks
   - Moderation Settings
   - Moderation Files

2. By default, team members can edit everything on the team site, so they could edit the staff list and questions. Lock the lists down as follows.

   Open the team's SharePoint site (step 1) and select **⚙ → Site contents**.

   **A. Make four lists read-only for teachers.** Do this for **Moderation Staff**, **Moderation Questions**, **Moderation Settings** and **Moderation Files**:
   1. Open the list, then select **⚙ → List settings** (for Moderation Files it's **Library settings**, or **More library settings**).
   2. Select **Permissions for this list → Stop Inheriting Permissions → OK**.
   3. Tick **science_Team Members → Edit User Permissions**, untick **Edit**, tick **Read** → **OK**.

   **B. Let teachers save marks.** On **Moderation Marks**:
   1. **⚙ → List settings → Permissions for this list → Stop Inheriting Permissions → OK**.
   2. Tick **science_Team Members → Edit User Permissions**, untick **Edit**, tick **Contribute** → **OK**.
   3. Back on **List settings**, select **Advanced settings → Item-level permissions** and set:
      - Read access: **Read all items**
      - Create and Edit access: **Create items and edit items that were created by the user**

      Select **OK**.

   Team owners keep full control of all five lists, so you and the Heads of Subject can still change everything.

## 8. Add your staff

1. In the app, open **Staff & settings → Add staff member**. For each person enter:
   - name;
   - **the email address they sign in to Teams with**;
   - role;
   - key stages;
   - subjects.
2. Make sure each person is in science_Team with the matching role (step 1):
   - CC and Heads of Subject → **Owner**
   - Teachers → **Member**

   Anyone added to the team later gets the right access automatically.
3. Optional: **Load sample team** shows how the dashboards look. **Remove sample data** deletes it.

Done. Staff open the tab and are signed in automatically.

---

## Problems and fixes

| What you see | Fix |
|---|---|
| "Setup isn't finished" | config.js still has placeholder values (step 5). |
| Sign-in error mentioning **AADSTS50011** or a redirect | The two addresses in step 3 and step 4 must match the setup helper exactly. |
| "This account can't open the SharePoint site" | Check `siteUrl` in config.js stops at the site name (step 1), and that the person is in science_Team. |
| "SharePoint refused this change" | Heads of Subject must be team **Owners**. Teachers need **Contribute** on Moderation Marks (step 7B). |
| A teacher can edit questions or staff | Step 7A hasn't been done for that list. |
| A teacher sees "You're not on the staff list yet" | The email in the staff list must match the address they sign in with; the screen shows the one to use. |
| The tab is blank or "refused to connect" | Re-download the package from the setup helper and upload it again. This is needed if the app's web address changed. |
| My change to config.js hasn't taken effect | GitHub Pages takes 1–2 minutes to update. Then reload the tab (in Teams: the tab's **⋯ → Reload**). |

## Updating the app later

- **New version of the code:** upload the replacement files to the same GitHub repository. Staff get it next time they open the tab, with no Teams changes needed.
- **When to re-upload to Teams:** only if the web address changes. Download the package from the setup helper with a higher version number (e.g. 1.0.1) and upload it again in the Teams admin centre.
