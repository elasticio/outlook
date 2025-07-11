/* eslint-disable no-restricted-syntax */
const { axiosReqWithRetryOnServerError } = require('@elastic.io/component-commons-library/dist/src/externalApi');
const { version, dependencies } = require('@elastic.io/component-commons-library/package.json');
const commons = require('@elastic.io/component-commons-library');
const compJson = require('../../component.json');
const packageJson = require('../../package.json');

const auth = {
  username: process.env.ELASTICIO_API_USERNAME,
  password: process.env.ELASTICIO_API_KEY,
};

const headers = { 'User-Agent': `${compJson.title}/${compJson.version} component-commons-library/${version} axios/${dependencies.axios}` };
const secretsUrl = `${process.env.ELASTICIO_API_URI}/v2/workspaces/${process.env.ELASTICIO_WORKSPACE_ID}/secrets/`;

module.exports.getSecret = async function getSecret(credentialId) {
  try {
    const response = await axiosReqWithRetryOnServerError.call(this, {
      method: 'GET',
      url: `${secretsUrl}${credentialId}`,
      auth,
      headers,
    });
    const { attributes } = response.data.data;
    return { ...attributes };
  } catch (err) {
    this.logger.error('Error refreshing secret:', err);
    return null;
  }
};

module.exports.refreshSecret = async function refreshSecret(credentialId) {
  try {
    const response = await axiosReqWithRetryOnServerError.call(this, {
      method: 'POST',
      url: `${secretsUrl}${credentialId}/refresh`,
      auth,
      headers,
    });
    return response.data.data.attributes;
  } catch (err) {
    this.logger.error('Error refreshing secret:', err);
    return null;
  }
};

module.exports.isNumberNaN = (num) => Number(num).toString() === 'NaN';

const getUserAgent = () => {
  const { version: compVersion } = compJson;
  const libVersion = packageJson.dependencies['@elastic.io/component-commons-library'];
  return `${compJson.title}/${compVersion} component-commons-library/${libVersion}`;
};

module.exports.buildMessage = async (msg) => {
  const {
    ccRecipients,
    saveToSentItems,
    subject,
    toRecipients,
    attachments: attachmentsIn,
  } = msg.body;
  const attachmentsOut = [];
  if (attachmentsIn && attachmentsIn.length > 0) {
    const attachmentsProcessor = new commons.AttachmentProcessor(getUserAgent(), msg.id);
    for (const attachment of attachmentsIn) {
      const { name, url } = attachment;
      if (!name || !url) throw new Error('Both "File name" and "URL to file" fields must be provided');
      const { data } = await attachmentsProcessor.getAttachment(url, 'arraybuffer');
      const attachmentBody = {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name,
        contentBytes: Buffer.from(data, 'binary').toString('base64'),
      };
      attachmentsOut.push(attachmentBody);
    }
  }
  const result = {
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
  if (attachmentsOut.length > 0) result.message.attachments = attachmentsOut;
  return result;
};

module.exports.getUserAgent = getUserAgent;
