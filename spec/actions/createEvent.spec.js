const { Logger } = require('@elastic.io/component-commons-library');
const nock = require('nock');
const chai = require('chai');
const sinon = require('sinon');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/createEvent');
const configuration = require('../data/configuration.new.in.json');
const jsonIn = require('../data/createEvent_test.in.json');
const jsonOut = require('../data/createEvent_test.out.json');

const cfgString = JSON.stringify(configuration);

describe('Outlook Create Event', () => {
  const refreshTokenUri = 'https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = `/me/calendars/${configuration.calendarId}/events`;

  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    self = {
      emit: sinon.spy(),
      logger,
    };
  });
  afterEach(() => {
    nock.cleanAll();
  });

  it('should emit (data and end events on success create request - case: http 200', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });
    const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
      .reply(200, jsonOut);

    await action.process.call(self, { body: jsonIn }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.be.eql(3);
    expect(args[0][0]).to.be.eql('updateKeys');
    expect(args[0][1].access_token).to.be.eql(1);
    expect(args[1][0]).to.be.eql('data');
    expect(args[2][0]).to.be.eql('end');
    expect(args[1][1].body).to.be.eql({
      id: 'testid12345',
      subject: 'Unit Test - Simple Event',
      body: {
        contentType: 'HTML',
        content: 'This is a test.',
      },
      start: {
        dateTime: '2017-05-19T18:00:00',
        timeZone: 'Central European Standard Time',
      },
      end: {
        dateTime: '2017-05-20T19:00:00',
        timeZone: 'Central European Standard Time',
      },
      calendarId: 'AAMkAGYyNmJlYjBmLTgwOWYtNGU0Mi04NW',
    });
    expect(scope1.isDone()).to.be.eql(true);
    expect(scope2.isDone()).to.be.eql(true);
  });

  it('should emit (data and end events on success create request - case: http 201', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });
    const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
      .reply(201, jsonOut);

    await action.process.call(self, { body: jsonIn }, cfg, {});

    const { callCount, args } = self.emit;
    expect(args[0][0]).to.eql('updateKeys');
    expect(args[0][1].access_token).to.eql(1);
    expect(args[1][0]).to.eql('data');
    expect(callCount).to.eql(3);
    expect(args[2][0]).to.eql('end');
    expect(args[1][1].body).to.eql({
      id: 'testid12345',
      subject: 'Unit Test - Simple Event',
      body: {
        contentType: 'HTML',
        content: 'This is a test.',
      },
      start: {
        dateTime: '2017-05-19T18:00:00',
        timeZone: 'Central European Standard Time',
      },
      end: {
        dateTime: '2017-05-20T19:00:00',
        timeZone: 'Central European Standard Time',
      },
      calendarId: 'AAMkAGYyNmJlYjBmLTgwOWYtNGU0Mi04NW',
    });
    expect(scope1.isDone()).to.be.eql(true);
    expect(scope2.isDone()).to.be.eql(true);
  });

  it('should emit error and end events on unsuccessful refresh token request', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(401, {
        access_token: 1,
      });
    await action.process.call(self, {
      body: jsonIn,
    }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(2);
    expect(args[0][0]).to.eql('error');
    expect(args[1][0]).to.eql('end');
    expect(scope1.isDone()).to.eql(true);
  });

  it('should emit error and end events on unsuccessful create request - case: bad request', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
      .reply(400, {});

    await action.process.call(self, {
      body: jsonIn,
    }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(3);
    expect(args[0][0]).to.eql('updateKeys');
    expect(args[0][1].access_token).to.eql(1);
    expect(args[1][0]).to.eql('error');
    expect(args[2][0]).to.eql('end');
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
  });

  it('should emit error and end events on unsuccessful create request - case: consent problems', async () => {
    const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
      .reply(200, {
        access_token: 1,
        expires_in: 3600,
      });

    const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
      .reply(403, jsonOut);

    await action.process.call(self, {
      body: jsonIn,
    }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(3);
    expect(args[0][0]).to.eql('updateKeys');
    expect(args[0][1].access_token).to.eql(1);
    expect(args[1][0]).to.eql('error');
    expect(args[2][0]).to.eql('end');
    expect(scope1.isDone()).to.eql(true);
    expect(scope2.isDone()).to.eql(true);
  });
});
