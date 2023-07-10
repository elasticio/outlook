const { Client } = require('./client');

const MAIL_RETRIEVE_MAX_COUNT = process.env.MAIL_RETRIEVE_MAX_COUNT || 1000;
const TOP_LIST_MAIL_FOLDER = process.env.TOP_LIST_MAIL_FOLDER || 100;

class OutlookClient extends Client {
  async getUserInfo() {
    return this.apiRequest({
      url: '/me',
      method: 'GET',
    });
  }

  async sendMail() {
    const data = {
      message: {
        subject: 'MS shit?',
        body: {
          contentType: 'Text',
          content: 'MS shit!',
        },
        toRecipients: [
          {
            emailAddress: {
              address: 'voropaiev@gmail.com',
            },
          },
        ],
      },
      saveToSentItems: false,
    };
    const response = await this.apiRequest({
      url: '/me/sendMail',
      method: 'POST',
      data,
    });
    return response.value;
  }

  async getMailFolders() {
    const response = await this.apiRequest({
      url: `/me/mailFolders?$top=${TOP_LIST_MAIL_FOLDER}`,
      method: 'GET',
    });
    return response.value;
  }

  async listChildFolders(parentId) {
    const response = await this.apiRequest({
      url: `/me/mailFolders/${parentId}/childFolders?$top=${TOP_LIST_MAIL_FOLDER}`,
      method: 'GET',
    });
    return response.value;
  }

  async getLatestMessages(folderId, lastModifiedDateTime, pollOnlyUnreadMail) {
    const filter = `$filter=lastModifiedDateTime gt ${lastModifiedDateTime}${pollOnlyUnreadMail ? ' and isRead eq false' : ''}`;
    const url = `/me/mailFolders/${folderId}/messages?$orderby=lastModifiedDateTime asc&$top=${MAIL_RETRIEVE_MAX_COUNT}&${filter}`;
    const response = await this.apiRequest({
      url,
      method: 'GET',
    });
    return response.value;
  }

  async getMyLatestContacts(lastModifiedDateTime) {
    const response = await this.apiRequest({
      url: `/me/contacts?$orderby=lastModifiedDateTime asc&$top=900&$filter=lastModifiedDateTime gt ${lastModifiedDateTime}`,
      method: 'GET',
    });
    return response.value;
  }

  async getMyLatestEvents(currentTime) {
    const response = await this.apiRequest({
      url: `/me/events?$filter=start/dateTime le '${currentTime}' and end/dateTime ge '${currentTime}'`,
      method: 'GET',
    });
    return response.value;
  }

  async getMyCalendars() {
    const response = await this.apiRequest({
      url: '/me/calendars',
      method: 'GET',
    });
    return response.value;
  }

  async createEvent(calendarId, body) {
    return this.apiRequest({
      url: `/me/calendars/${calendarId}/events`,
      method: 'POST',
      body,
    });
  }

  async moveMessage(originalMailFolders, messageId, destinationId) {
    return this.apiRequest({
      url: `/me/mailFolders/${originalMailFolders}/messages/${messageId}/move`,
      method: 'POST',
      body: {
        destinationId,
      },
    });
  }

  async getDeletedItemsFolder() {
    return this.apiRequest({
      url: '/me/mailFolders/deleteditems',
      method: 'GET',
    });
  }

  async getMailFolderById(folderId) {
    return this.apiRequest({
      url: `/me/mailFolders/${folderId}`,
      method: 'GET',
    });
  }
}

module.exports.OutlookClient = OutlookClient;
