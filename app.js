const express = require("express");
const msal = require("@azure/msal-node");
const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");
require('dotenv').config();

const app = express();
const port = 3000;

// MSAL configuration
const msalConfig = {
  auth: {
    clientId: process.env.YOUR_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.YOUR_TENANT_ID}`,
    clientSecret: process.env.YOUR_CLIENT_SECRET,
  },
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

// Get access token
async function getAccessToken() {
  const authResult = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  console.log(authResult);
  return authResult.accessToken;
}

// Initialize Graph client
async function getGraphClient() {
  const accessToken = await getAccessToken();
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
  return client;
}

// async function createSubscription() {
//   const client = await getGraphClient();
//   const subscription = await client.api("/subscriptions").post({
//     changeType: "updated",
//     notificationUrl: "http://localhost:3000/notifications",
//     resource: "/me/drive/root",
//     expirationDateTime: new Date(Date.now() + 3600000).toISOString(), // 1 hour expiration
//     clientState: "secretClientValue",
//   });
//   console.log(subscription);
// }

// // Call createSubscription function to create the webhook subscription
// createSubscription();

// List files
app.get("/files", async (req, res) => {
  const client = await getGraphClient();
  const files = await client.api("/users/134/root/children").get();
  res.json();
});

// Download file
app.get("/files/download/:fileId", async (req, res) => {
  const fileId = req.params.fileId;
  const client = await getGraphClient();
  const file = await client.api(`/me/drive/items/${fileId}`).get();
  res.redirect(file["@microsoft.graph.downloadUrl"]);
});

// List users with access to a file
app.get("/files/:fileId/permissions", async (req, res) => {
  const fileId = req.params.fileId;
  const client = await getGraphClient();
  const permissions = await client
    .api(`/me/drive/items/${fileId}/permissions`)
    .get();
  res.json(permissions);
});

app.post("/notifications", (req, res) => {
  // Verify the validation token sent by Microsoft Graph
  if (req.query.validationToken) {
    res.send(req.query.validationToken);
  } else {
    // Handle the notification
    console.log(req.body);
    res.sendStatus(202);
  }
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
