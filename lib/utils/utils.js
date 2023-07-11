const { axiosReqWithRetryOnServerError } = require('@elastic.io/component-commons-library/dist/src/externalApi');
const { version, dependencies } = require('@elastic.io/component-commons-library/package.json');
const compJson = require('../../component.json');

const auth = {
  username: process.env.ELASTICIO_API_USERNAME,
  password: process.env.ELASTICIO_API_KEY,
};

const headers = { 'User-Agent': `${compJson.title}/${compJson.version} component-commons-library/${version} axios/${dependencies.axios}` };
const secretsUrl = `${process.env.ELASTICIO_API_URI}/v2/workspaces/${process.env.ELASTICIO_WORKSPACE_ID}/secrets/`;

module.exports.getSecret = async (credentialId) => {
  const response = await axiosReqWithRetryOnServerError.call(this, {
    method: 'GET',
    url: `${secretsUrl}${credentialId}`,
    auth,
    headers,
  });
  const { attributes } = response.data.data;
  return { ...attributes };
};

module.exports.refreshSecret = async (credentialId) => {
  const response = await axiosReqWithRetryOnServerError.call(this, {
    method: 'POST',
    url: `${secretsUrl}${credentialId}/refresh`,
    auth,
    headers,
  });
  return response.data.data.attributes;
};

module.exports.isNumberNaN = (num) => Number(num).toString() === 'NaN';

module.exports.buildMessage = (msg) => {
  const {
    ccRecipients,
    saveToSentItems,
    subject,
    toRecipients,
  } = msg.body;
  return {
    message: {
      ccRecipients,
      body: {
        contentType: msg.body.body.contentType || 'text',
        content: msg.body.body.content,
      },
      subject,
      toRecipients,
    },
    saveToSentItems: saveToSentItems || true,
  };
};
