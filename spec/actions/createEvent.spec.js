'use strict';
describe('Outlook Create Event', function () {

  const nock = require('nock');
  const action = require('../../lib/actions/createEvent');

  const cfg = require('../data/configuration.in.json');
  const jsonIn = require('../data/createEvent_test.in.json');
  const jsonOut = require('../data/createEvent_test.out.json');

  const refreshTokenUri ='https://login.microsoftonline.com';
  const refreshTokenApi = '/common/oauth2/v2.0/token';
  const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
  const microsoftGraphApi = `/me/calendars/${cfg.calendarId}/events`;


  var self;
  beforeEach(function() {
    self = jasmine.createSpyObj('self', ['emit']);
  });

  it('should emit (data and end events on success create request - case: http 200', function (done) {
    const scope1 = nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1});

    const scope2 = nock(microsoftGraphUri)
      .post(microsoftGraphApi)
      .reply(200, jsonOut);

    action.process.call(self, {body: jsonIn}, cfg, {})
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults() {
      let calls = self.emit.calls;
      expect(calls.length).toEqual(2);
      expect(calls[0].args[0]).toEqual('data');
      expect(calls[1].args[0]).toEqual('end');
      expect(calls[0].args[1].body).toEqual({ "id" : "testid12345", "calendarId" : "AAMkAGYyNmJlYjBmLTgwOWYtNGU0Mi04NW" });
      expect(scope1).toBeTruthy();
      expect(scope2).toBeTruthy();
    };
  });

  it('should emit (data and end events on success create request - case: http 201', function (done) {
    const scope1 = nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1})

    const scope2 = nock(microsoftGraphUri)
      .post(microsoftGraphApi)
      .reply(201, jsonOut);

    action.process.call(self, {body: jsonIn}, cfg, {})
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults() {
      let calls = self.emit.calls;
      expect(calls.length).toEqual(2);
      expect(calls[0].args[0]).toEqual('data');
      expect(calls[1].args[0]).toEqual('end');
      expect(calls[0].args[1].body).toEqual({"id": "testid12345", "calendarId": "AAMkAGYyNmJlYjBmLTgwOWYtNGU0Mi04NW"});
      expect(scope1).toBeTruthy();
      expect(scope2).toBeTruthy();
    };
  });

  it('should emit error and end events on unsuccessful refresh token request', function (done) {
    const scope1 = nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(401, {access_token: 1});

    action.process.call(self, {body: jsonIn}, cfg, {})
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults() {
      let calls = self.emit.calls;
      expect(calls.length).toEqual(2);
      expect(calls[0].args[0]).toEqual('error');
      expect(calls[1].args[0]).toEqual('end');
      expect(scope1).toBeTruthy();
    };
  });

  it('should emit error and end events on unsuccessful create request - case: bad request', function (done) {
    const scope1 = nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1});

    const scope2 = nock(microsoftGraphUri)
      .post(microsoftGraphApi)
      .reply(400, jsonOut);

    action.process.call(self, {body: jsonIn}, cfg, {})
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults() {
      let calls = self.emit.calls;
      expect(calls.length).toEqual(2);
      expect(calls[0].args[0]).toEqual('error');
      expect(calls[1].args[0]).toEqual('end');
      expect(scope1).toBeTruthy();
      expect(scope2).toBeTruthy();
    };
  });

  it('should emit error and end events on unsuccessful create request - case: consent problems', function (done) {
    const scope1 = nock(refreshTokenUri)
      .post(refreshTokenApi)
      .reply(200, {access_token: 1});

    const scope2 = nock(microsoftGraphUri)
      .post(microsoftGraphApi)
      .reply(403, jsonOut);

    action.process.call(self, {body: jsonIn}, cfg, {})
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults() {
      let calls = self.emit.calls;
      expect(calls.length).toEqual(2);
      expect(calls[0].args[0]).toEqual('error');
      expect(calls[1].args[0]).toEqual('end');
      expect(scope1).toBeTruthy();
      expect(scope2).toBeTruthy();
    };
  });

});
