const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('dotenv').config();

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/createEvent');

describe('Outlook Create Event', () => {
  let self;
  let cfg;
  beforeEach(() => {
    self = {
      emit: sinon.spy(),
      logger,
    };
    cfg = {
      oauth2: {
        token_type: 'Bearer',
        scope: 'openid User.Read Contacts.Read profile Calendars.ReadWrite Mail.ReadWrite Files.ReadWrite.All',
        expires_in: 3600,
        ext_expires_in: 3600,
        tokenExpiryTime: (new Date('2995-12-17T03:24:00')).toISOString(),
        client_id: process.env.OAUTH_CLIENT_ID,
        client_secret: process.env.OAUTH_CLIENT_SECRET,
        access_token: process.env.ACCESS_TOKEN,
      },
      folderId: process.env.FOLDER_ID,
      calendarId: 'AQMkADAwATM0MDAAMS0yMgA3MC1mZDkyLTAwAi0wMAoARgAAA14NR8zJ7GxJnz_JauTU1uQHALdeJHYCtWNEt1wFU6jUUYgAAAIBBgAAALdeJHYCtWNEt1wFU6jUUYgAAAIhdQAAAA==',
      timeZone: 'Europe/Kiev',
      importance: 'Normal',
      showAs: 'Busy',
    };
  });

  it('getCalendars test', async () => {
    const result = await action.getCalendars.call(self, cfg);
    expect(result).to.not.eql({});
  });

  it('Process Action', async () => {
    const msg = {
      body: {
        subject: 'PS Test Event',
        body: {
          content: 'PS Test Event',
        },
        start: {
          dateTime: '2023-07-18T10:00:00.0000000',
        },
        end: {
          dateTime: '2023-07-18T12:00:00.0000000',
        },
      },
    };
    await action.process.call(self, msg, cfg);
  });
});
