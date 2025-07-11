/* eslint-disable no-restricted-syntax, class-methods-use-this, no-param-reassign */
const {
  axiosReqWithRetryOnServerError,
  getErrMsg,
  sleep,
} = require('@elastic.io/component-commons-library');
const {
  getSecret,
  isNumberNaN,
  refreshSecret,
} = require('./utils/utils');

class Client {
  constructor(context, cfg) {
    this.logger = context.logger;
    this.cfg = cfg;
    this.retries = 5;
  }

  setLogger(logger) { this.logger = logger; }

  async apiRequest(opts) {
    if (!this.accessToken) {
      this.logger.info('Token not found, going to fetch new one');
      await this.getNewAccessToken();
      this.logger.info('Token created successfully');
    }

    const baseURL = 'https://graph.microsoft.com/v1.0/';

    opts = {
      ...opts,
      baseURL,
      headers: {
        ...opts.headers || {},
        Authorization: `Bearer ${this.accessToken}`,
      },
    };

    let response;
    let error;
    let currentRetry = 0;

    while (currentRetry < this.retries) {
      try {
        response = await axiosReqWithRetryOnServerError.call(this, opts);
        return response;
      } catch (err) {
        error = err;
        if (err.response) this.logger.error(getErrMsg(err.response));
        if (err.response?.status === 401) {
          this.logger.info('Token invalid, going to fetch new one');
          const currentToken = this.accessToken;
          await this.getNewAccessToken();
          if (currentToken === this.accessToken) {
            this.logger.info('Token not changed, going to force refresh');
            await this.refreshAndGetNewAccessToken();
          }
          this.logger.info('Trying to use new token');
          opts.headers.Authorization = `Bearer ${this.accessToken}`;
        } else if (err.response?.status === 429) {
          const retryAfter = 2 ** (currentRetry + 1);
          this.logger.error(`Going to retry after ${retryAfter}sec (${currentRetry + 1} of ${this.retries})`);
          await sleep(retryAfter * 1000);
        } else {
          throw err;
        }
      }
      currentRetry++;
    }
    throw error;
  }

  async getNewAccessToken() {
    let fields;
    if (this.cfg.secretId) {
      this.logger.info('Fetching credentials by secretId');
      const response = await getSecret.call(this, this.cfg.secretId);
      this.accessToken = response.credentials.access_token;
      fields = response.credentials.fields;
    } else if (this.cfg.oauth) {
      this.logger.info('Fetching credentials from this.cfg');
      this.accessToken = this.cfg.oauth.access_token;
      fields = this.cfg;
    } else {
      throw new Error('Can\'t find credentials in incoming configuration');
    }

    const { retries = this.retries } = fields || {};
    if (!isNumberNaN(retries) && Number(retries) <= 10) this.retries = Number(retries);
  }

  async refreshAndGetNewAccessToken() {
    const response = await refreshSecret.call(this, this.cfg.secretId);
    this.accessToken = response.credentials.access_token;
  }
}

module.exports.Client = Client;
