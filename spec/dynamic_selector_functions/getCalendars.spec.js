'use strict'
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

  /*
  it('should return calendar info on success get request', function (done) {
    const scope1 = nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1});

    const scope2 = nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, jsonOut);

    action.getCalendars(cfg, cb)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults() {
      expect(cb).toHaveBeenCalled();
      expect(cb.calls.length).toEqual(1);
      expect(cb).toHaveBeenCalledWith(null, { "AAMkAGI2TGuLAAA=" : "Calendar" });
      expect(scope1).toBeTruthy();
      expect(scope2).toBeTruthy();
    }
  });

  it('should return errors on refresh token failure ', function (done) {
    const scope1 = nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(401, {access_token: 1});

    const scope2 = nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, jsonOut);

    action.getCalendars(cfg, cb)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults() {
      expect(cb).toHaveBeenCalled();
      expect(cb.calls.length).toEqual(1);
      expect(cb).toHaveBeenCalledWith({ statusCode : 401 });
      expect(scope1).toBeTruthy();
      expect(scope2).toBeTruthy();
    }
  });
  */

});
