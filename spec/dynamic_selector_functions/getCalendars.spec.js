'use strict';
describe('Outlook Get Calendars', function () {
  const nock = require('nock');
  const action = require('../../lib/actions/createEvent');
  const Q = require('q');

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

  it('should return calendar info on success get request', function (done) {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1});

    nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, jsonOut);

    Q.ninvoke(action, "getCalendars", cfg)
      .then(checkResults)
      .then(done)
      .catch(done.fail)

    function checkResults(data) {
      expect(data).toEqual({ 'AAMkAGI2TGuLAAA=' : 'Calendar' });
    }
  });

  it('should return errors on refresh token failure ', function (done) {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(401, {access_token: 1});

    nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, jsonOut);

    Q.ninvoke(action, "getCalendars", cfg)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(data) {
      expect(data).toBeUndefined();
    }

    function checkError(err) {
      expect('StatusCodeError').toEqual(err.name);
      expect(401).toEqual(err.statusCode);
    }

  });


});
