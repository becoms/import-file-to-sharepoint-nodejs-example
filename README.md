# Upload a file in a Sharepoint using Node.js

This sample NodeJS web application shows how to create a file in a sharepoint, using :

- node-js
- typescript

To communicate with the sharepoint, we use :

- @azure/msal-node
- @microsoft/microsoft-graph-client

## Table of Contents

- [Using the Sample](#using-the-sample)
- [Steps to put your file](#steps-to-put-your-file)

## Using the Sample

### Prerequisites

To use the sample, you need the following:

- [Node.js](https://nodejs.org/) version 18 or 20.
- A [work or school account](https://developer.microsoft.com/microsoft-365/dev-program).
- The application ID and key from the application that you register on the Azure Portal.
- The permission to write in your sharepoint

### Configure and run the sample

1. Rename [.env.defaults](.env.defaults) to **.env** and open it in a text editor.

1. Replace `SHAREPOINT_CLIENT_ID` with the client ID of your registered Azure application.

1. Replace `SHAREPOINT_CLIENT_SECRET` with the client secret of your registered Azure application.

1. Replace `SHAREPOINT_TENANT_ID` with the tenant ID of your organization. This information can be found next to the client ID on the application management page, note: if you choose _Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)_ replace this value for "common".

1. Replace `SHAREPOINT_URL` with the url of your sharepoint (must look like `yourSharepointName.sharepoint.com`).

1. Replace `SHAREPOINT_SUBSITE` with the name of your sharepoint site (if your sharepoint URL is 'https://yourSharepointName.sharepoint.com/sites/yourSharepointSubsite', it must be `yourSharepointSubsite`).

1. Replace `SHAREPOINT_PATH` with the path of the directory you want to put your file.

1. Install the dependencies running the following command:

   ```bash
   yarn
   ```

1. Build the application with the following command:

   ```bash
   yarn build
   ```

1. Start the application with the following command:

   ```bash
   yarn start
   ```

1. Open a browser / client and go to [http://localhost:8080](http://localhost:8080). This will create an excel file named `#test.xlsx` at the path you wrote in your .env

## Steps to upload your file

### Get access token

You need an access token and a drive id to build your API URL that allows you to put a file in your sharepoint.
Here is how to get your access token, that will be necessary in every next steps.

```
import msal from "@azure/msal-node";

const cca = new msal.ConfidentialClientApplication({
  auth: {
    clientId: process.env.SHAREPOINT_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.SHAREPOINT_TENANT_ID}`,
    clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
  },
});

const getAccessToken = async () => {
  const result = await cca.acquireTokenByClientCredential({
    scopes: [`https://graph.microsoft.com/.default`],
  });

  return result?.accessToken;
  };
```

### (optional) Get graph client

We'll use the graph client in the following steps, but you can use `fetch`, `axios` or any client if you prefer.

```
import { Client } from "@microsoft/microsoft-graph-client";

const getGraphClient = async (accessToken: string) => {
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
  return client;
};
```

### Get your site id

You need an access token and a drive id to build your API URL that allows you to put a file in your sharepoint.
In order to get your drive id, you have to get your site id.

```
async function getSharePointSiteId({ client }: { client: Client }) {
  const apiUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_URL}:/sites/${process.env.SHAREPOINT_SUBSITE}`;

  try {
    const site = await client.api(apiUrl).get();

    var siteId = site.id.split(",")[1];
    return siteId;
  } catch (error) {
    console.error("Error when getting SharePoint site:", error);
  }
}
```

We split the _site.id_ because it first has the site URL and then the site id, seperated with ','.

### Get the drive id

```
async function getSharePointDriveId({
  client,
  siteId,
}: {
  client: Client;
  siteId: string;
}) {
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/Drives`;

  try {
    const drives = await client.api(siteUrl).get();
    return drives.value[0].id;
  } catch (error) {
    console.error("Error when getting SharePoint drive:", error);
  }
}
```

### Upload the file

Encode file name because of special characters (for example here "#") because the fileName goes in the API URL. You will hase a content-type error if the file name can't be read by the API.
Make sure you have the right file extension in the name as well.
Example for an `.xlslx` file named `#test` :

```
const fileName = `${encodeURIComponent("#test")}.xlsx`;
```

#### Convert file to buffer

Convert an existing file (with a relative path) :

```
const file = fs.readFileSync("yourFilePath");
```

In our sample, we convert our worksheet (created from exceljs) to buffer :

```
const file = await workbook.xlsx.writeBuffer();
```

#### Put the file (buffer) in your sharepoint

If you use the graph client, Content-Type has to be specified, as the default is json in the graph client.

```
export const sharepointWriter = async () => {
  try {
    const accessToken = await getAccessToken();

    const client = await getGraphClient(accessToken);

    const siteId: string = await getSharePointSiteId({ client });

    const driveId = await getSharePointDriveId({
      client: client,
      siteId: siteId,
    });

    const siteUrl = `https://graph.microsoft.com/v1.0/Drives/${driveId}/root:/${process.env.SHAREPOINT_PATH}/${fileName}:/content`;

    //#region : Put file with microsoft graph api
    const response = await client
      .api(siteUrl)
      .header("Content-Type", "application/octet-stream")
      .put(file);

    return response;

    //#endregion : Put file with microsoft graph api
  } catch (e) {
    console.error(e);
  }
};
```

If you don't have direct access to the sharepoint, you can log the response and click on the _@microsoft.graph.downloadUrl_ link, and check the file

## Contributing
