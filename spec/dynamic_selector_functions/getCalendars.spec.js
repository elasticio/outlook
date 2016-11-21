'use strict'
describe('Outlook Get Calendars', function () {
  const nock = require('nock');
  const action = require('../../lib/actions/createEvent');

  const cfg = require('../data/configuration.in.json');

  const refreshTokenUri ='https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = '/me/calendars';

  const getCalendarsMockResult = { calendar: "testCalendar" };

  var cb;
  beforeEach(function () {
   cb = jasmine.createSpy('cb');
  });

  it('should return no errors on success create request', function () {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1})

    nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(200, getCalendarsMockResult);

    action.getCalendars(cfg, cb);

    waitsFor(function () {
      return cb.callCount;
    });

    runs(function () {
      expect(cb).toHaveBeenCalled();
      expect(cb.calls.length).toEqual(1);
      expect(cb).toHaveBeenCalledWith(null, { });
   });
  });

  it('should return errors on refresh token failure', function () {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
        .reply(401, {access_token: 1})

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

  it('should return errors  on request failure: case - consent problems', function () {
    nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1})

    nock(microsoftGraphUri)
      .get(microsoftGraphApi)
      .reply(403, getCalendarsMockResult);

    action.getCalendars(cfg, cb);

    waitsFor(function () {
      return cb.callCount;
    });

    runs(function () {
      expect(cb).toHaveBeenCalled();
      expect(cb.calls.length).toEqual(1);
      expect(cb).toHaveBeenCalledWith({ statusCode : 403 });
    });
  });

});
