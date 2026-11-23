/* Moderation Hub settings. Edit the four values marked EDIT, then save.
   The setup guide (SETUP.md) says where each one comes from. */
window.CONFIG = {
  // EDIT 1: Entra admin centre > App registrations > Moderation Hub > Overview > "Application (client) ID"
  clientId: "PASTE-APPLICATION-CLIENT-ID",

  // EDIT 2: same Overview page > "Directory (tenant) ID"
  tenantId: "PASTE-DIRECTORY-TENANT-ID",

  // EDIT 3: the SharePoint site you created, exactly as it appears in the browser address bar (no trailing pages)
  siteUrl: "https://YOURSCHOOL.sharepoint.com/sites/ScienceModeration",

  // EDIT 4: curriculum coordinator email(s). These people always have full access.
  adminEmails: ["cc@yourschool.org.uk"],

  // Optional: shown on the app's details page in Teams
  schoolName: "Our School",

  // Do not change. Teams uses this to recognise the app when you update it.
  appId: "3e1a863d-7b76-4e6a-9ff3-2f50dc53b975"
};
