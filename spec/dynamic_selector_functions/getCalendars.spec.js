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

  it('should return no errors on success create request', function () {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1})

    nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, jsonOut);

    var resultsPromise = action.getCalendars(cfg);

    var result;
    runs(function () {
      resultsPromise.then(function (data) {
        result = data;
      });
    });

    waitsFor(function () {
      return result;
    });

    runs(function () {
      expect(result).toBeDefined();
      expect(result).toEqual( { "AAMkAGI2TGuLAAA=": "Calendar" });
    });

   });

});
