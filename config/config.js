const msalConfig = {
    auth: {
      clientId: 'YOUR_CLIENT_ID',
      authority: `https://login.microsoftonline.com/YOUR_TENANT_ID`,
      clientSecret: 'YOUR_CLIENT_SECRET'
    }
  };

module.exports = msalConfig;