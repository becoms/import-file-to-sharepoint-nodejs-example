import msal from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { Buffer } from "exceljs";

const path = process.env.SHAREPOINT_PATH;
const siteName = process.env.SHAREPOINT_URL;
const subsiteName = process.env.SHAREPOINT_SUBSITE;
const clientId = process.env.SHAREPOINT_CLIENT_ID;
const clientSecret = process.env.SHAREPOINT_CLIENT_SECRET;
const tenantId = process.env.SHAREPOINT_TENANT_ID;

const scope = "https://graph.microsoft.com";

// Auth infos
const cca = new msal.ConfidentialClientApplication({
  auth: {
    clientId: clientId || "",
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret: clientSecret,
  },
});

// Function to get token to access to microsoft graph
const getAccessToken = async () => {
  const result = await cca.acquireTokenByClientCredential({
    scopes: [`${scope}/.default`],
  });

  console.debug("access token : ", result?.accessToken);
  return result?.accessToken;
};

// Function to get microsoft graph client, token needed
const getGraphClient = async (accessToken: string) => {
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
  return client;
};

// Function to get sharepoint site, token needed
// Returns site id
async function getSharePointSiteId({ client }: { client: Client }) {
  const apiUrl = `${scope}/v1.0/sites/${siteName}:/sites/${subsiteName}`;

  try {
    const site = await client.api(apiUrl).get();
    // We split the site.id because it first has the site URL and then the site id, seperated with ','
    var siteId = site.id.split(",")[1];
    return siteId;
  } catch (error) {
    console.error("Error when getting SharePoint site:", error);
  }
}

// Function to get sharepoint drives, site ID and token needed
// Returns drive id
async function getSharePointDriveId({
  client,
  siteId,
}: {
  client: Client;
  siteId: string;
}) {
  const siteUrl = `${scope}/v1.0/sites/${siteId}/Drives`;

  try {
    const drives = await client.api(siteUrl).get();
    return drives.value[0].id;
  } catch (error) {
    console.error("Error when getting SharePoint drive:", error);
  }
}

// Function to write in sharepoint, drive ID and token needed
export const sharepointWriter = async (
  file: Buffer | undefined,
  fileName: string
) => {
  try {
    const accessToken = await getAccessToken();
    if (!accessToken) return;
    const client = await getGraphClient(accessToken);

    const siteId: string = await getSharePointSiteId({ client });

    const driveId = await getSharePointDriveId({
      client: client,
      siteId: siteId,
    });

    const siteUrl = `${scope}/v1.0/Drives/${driveId}/root:/${path}/${fileName}:/content`;

    //#region : Put file with microsoft graph api
    // Content-Type here has to be specified, as the default is json.
    const response = await client
      .api(siteUrl)
      .header("Content-Type", "application/octet-stream")
      .put(file);

    // If you don't have direct access to the sharepoint, you can log the response and click on the "@microsoft.graph.downloadUrl" link, and check the file
    return response;

    //#endregion : Put file with microsoft graph api
  } catch (e) {
    console.error(e);
  }
};
