'use strict';
describe('Outlook Get Calendars', function () {
  const nock = require('nock');
  const action = require('../../lib/actions/createEvent');

  const cfg = require('../data/configuration.in.json');
  const jsonOut = require('../data/getCalendars_test.out.json');

  const refreshTokenUri ='https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = '/me/calendars';

  var cb;
  beforeEach(function () {
    cb = jasmine.createSpy('cb');
  });

  it('should return calendar info on success get request', function () {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1});

    nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, jsonOut);

    action.getCalendars(cfg, cb);

    waitsFor(function () {
      return cb.callCount;
    });

    runs(function () {
      expect(cb).toHaveBeenCalled();
      expect(cb.calls.length).toEqual(1);
      expect(cb).toHaveBeenCalledWith(null, { 'AAMkAGI2TGuLAAA=' : 'Calendar' });
    });
  });

  it('should return errors on refresh token failure ', function () {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(401, {access_token: 1});

    nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, jsonOut);

    action.getCalendars(cfg, cb);

    waitsFor(function () {
      return cb.callCount;
    });

    runs(function () {
      expect(cb).toHaveBeenCalled();
      expect(cb.calls.length).toEqual(1);
      expect(cb).toHaveBeenCalledWith({ statusCode : 401 });
    });
  });

});
