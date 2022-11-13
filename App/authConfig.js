const { LogLevel } = require("@azure/msal-node");
const { CLIENT_ID, TENANT_ID } = require("./config");

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/" + TENANT_ID,
  },
  system: {
    loggerOptions: {
      loggerCallback(_, message) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Verbose,
    },
  },
};

const protectedResources = {
  graphMe: {
    endpoint: "https://graph.microsoft.com/v1.0/me",
    scopes: ["User.Read"],
  },
};

module.exports = {
  msalConfig: msalConfig,
  protectedResources: protectedResources,
};
