const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

/**
 * Creating a Graph client instance via options method. For more information, visit:
 * https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/dev/docs/CreatingClientInstance.md#2-create-with-options
 * @param {String} accessToken
 * @returns
 */
const getGraphClient = (accessToken) => {
  const graphClient = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
  return graphClient;
};

module.exports = getGraphClient;
