const { LogLevel } = require("@azure/msal-node");

const AAD_ENDPOINT_HOST = "https://login.microsoftonline.com/"; // include the trailing slash
const TENANT_ID = "";
const msalConfig = {
  auth: {
    clientId: "Enter_the_Application_Id_Here",
    authority: AAD_ENDPOINT_HOST + TENANT_ID,
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

const GRAPH_ENDPOINT_HOST = "https://graph.microsoft.com/"; // include the trailing slash
const protectedResources = {
  graphMe: {
    endpoint: `${GRAPH_ENDPOINT_HOST}v1.0/me`,
    scopes: ["User.Read"],
  },
};

module.exports = {
  msalConfig: msalConfig,
  protectedResources: protectedResources,
};
