/** Single source of Graph scopes for login and token acquisition. */
export const GRAPH_SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Sites.ReadWrite.All",
];

/** Scopes for SharePoint Online REST (attachments, etc.). */
export const SHAREPOINT_SCOPES = [
  "https://adeptservicesaustralia.sharepoint.com/Sites.ReadWrite.All",
];
