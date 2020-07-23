const { OAuth2RestClient } = require('@elastic.io/component-commons-library');

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

  async getMessages(folderId) {
    const response = await this.restClient.makeRequest({
      url: `/me/mailFolders/${folderId}/messages`,
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
}

module.exports.Client = Client;
