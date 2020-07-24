const { OAuth2RestClient } = require('@elastic.io/component-commons-library');

const MAIL_RETRIEVE_MAX_COUNT = process.env.MAIL_RETRIEVE_MAX_COUNT || 1000;

class Client {
  constructor(emitter, cfg) {
    this.logger = emitter.logger;
    this.cfg = cfg;
    this.cfg.resourceServerUrl = 'https://graph.microsoft.com/v1.0';
    this.cfg.authorizationServerTokenEndpointUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
    this.cfg.oauth2_field_client_id = process.env.OAUTH_CLIENT_ID;
    this.cfg.oauth2_field_client_secret = process.env.OAUTH_CLIENT_SECRET;
    this.restClient = new OAuth2RestClient({ logger: this.logger, emit: emitter.emit }, this.cfg);
  }

  async getUserInfo() {
    return this.restClient.makeRequest({
      url: '/me',
      method: 'GET',
    });
  }

  async getMailFolders() {
    const response = await this.restClient.makeRequest({
      url: '/me/mailFolders',
      method: 'GET',
    });

    return response.value;
  }

  async getLatestMessages(folderId, lastModifiedDateTime) {
    const response = await this.restClient.makeRequest({
      url: `/me/mailFolders/${folderId}/messages?$orderby=lastModifiedDateTime asc&$top=${MAIL_RETRIEVE_MAX_COUNT}&$filter=lastModifiedDateTime gt ${lastModifiedDateTime}`,
      method: 'GET',
    });

    return response.value;
  }

  async getMyLatestContacts(lastModifiedDateTime) {
    const response = await this.restClient.makeRequest({
      url: `/me/contacts?$orderby=lastModifiedDateTime asc&$top=900&$filter=lastModifiedDateTime gt ${lastModifiedDateTime}`,
      method: 'GET',
    });

    return response.value;
  }

  async getMyLatestEvents(currentTime) {
    const response = await this.restClient.makeRequest({
      url: `/me/events?$filter=start/dateTime le '${currentTime}' and end/dateTime ge '${currentTime}'`,
      method: 'GET',
    });
    return response.value;
  }

  async getMyCalendars() {
    const response = await this.restClient.makeRequest({
      url: '/me/calendars',
      method: 'GET',
    });

    return response.value;
  }

  async createEvent(calendarId, body) {
    return this.restClient.makeRequest({
      url: `/me/calendars/${calendarId}/events`,
      method: 'POST',
      body,
    });
  }

  async moveMessage(originalMailFolders, messageId, destinationId) {
    return this.restClient.makeRequest({
      url: `/me/mailFolders/${originalMailFolders}/messages/${messageId}/move`,
      method: 'POST',
      body: {
        destinationId,
      },
    });
  }

  async getDeleteditemsFolder() {
    return this.restClient.makeRequest({
      url: '/me/mailFolders/deleteditems',
      method: 'GET',
    });
  }
}

module.exports.Client = Client;
