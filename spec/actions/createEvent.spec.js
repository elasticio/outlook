

describe('Outlook Create Event', function test() {

    const nock = require('nock');
    const action = require('../../lib/actions/createEvent');

    const cfg = require('../data/configuration.in.json');
    const jsonIn = require('../data/createEvent_test.in.json');
    const jsonOut = require('../data/createEvent_test.out.json');

    const refreshTokenUri = 'https://login.microsoftonline.com';
    const refreshTokenApi = '/common/oauth2/v2.0/token';
    const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
    const microsoftGraphApi = `/me/calendars/${cfg.calendarId}/events`;


    var self;
    beforeEach(function createSpy() {
        self = jasmine.createSpyObj('self', ['emit']);
    });

    it('should emit (data and end events on success create request - case: http 200', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(200, {
                access_token: 1
            });

        const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
            .reply(200, jsonOut);

        function checkResults() {
            let calls = self.emit.calls;
            expect(calls.count()).toEqual(2);
            expect(calls.argsFor(0)[0]).toEqual('data');
            expect(calls.argsFor(1)[0]).toEqual('end');
            expect(calls.argsFor(0)[1].body).toEqual({
                id: 'testid12345',
                subject: 'Unit Test - Simple Event',
                body: {
                    contentType: 'HTML',
                    content: 'This is a test.'
                },
                start: {
                    dateTime: '2017-05-19T18:00:00',
                    timeZone: 'Central European Standard Time'
                },
                end: {
                    dateTime: '2017-05-20T19:00:00',
                    timeZone: 'Central European Standard Time'
                },
                calendarId: 'AAMkAGYyNmJlYjBmLTgwOWYtNGU0Mi04NW'
            });
            expect(scope1.isDone()).toBeTruthy();
            expect(scope2.isDone()).toBeTruthy();
        }

        action.process.call(self, {
            body: jsonIn
        }, cfg, {})
            .then(checkResults)
            .then(done)
            .catch(done.fail);
    });

    it('should emit (data and end events on success create request - case: http 201', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(200, {
                access_token: 1
            });

        const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
            .reply(201, jsonOut);

        function checkResults() {
            let calls = self.emit.calls;
            expect(calls.count()).toEqual(2);
            expect(calls.argsFor(0)[0]).toEqual('data');
            expect(calls.argsFor(1)[0]).toEqual('end');
            expect(calls.argsFor(0)[1].body).toEqual({
                id: 'testid12345',
                subject: 'Unit Test - Simple Event',
                body: {
                    contentType: 'HTML',
                    content: 'This is a test.'
                },
                start: {
                    dateTime: '2017-05-19T18:00:00',
                    timeZone: 'Central European Standard Time'
                },
                end: {
                    dateTime: '2017-05-20T19:00:00',
                    timeZone: 'Central European Standard Time'
                },
                calendarId: 'AAMkAGYyNmJlYjBmLTgwOWYtNGU0Mi04NW'
            });
            expect(scope1.isDone()).toBeTruthy();
            expect(scope2.isDone()).toBeTruthy();
        }

        action.process.call(self, {
            body: jsonIn
        }, cfg, {})
            .then(checkResults)
            .then(done)
            .catch(done.fail);
    });

    it('should emit error and end events on unsuccessful refresh token request', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(401, {
                access_token: 1
            });

        function checkResults() {
            let calls = self.emit.calls;
            expect(calls.count()).toEqual(2);
            expect(calls.argsFor(0)[0]).toEqual('error');
            expect(calls.argsFor(1)[0]).toEqual('end');
            expect(scope1.isDone()).toBeTruthy();
        }

        action.process.call(self, {
            body: jsonIn
        }, cfg, {})
            .then(checkResults)
            .then(done)
            .catch(done.fail);

    });

    it('should emit error and end events on unsuccessful create request - case: bad request', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(200, {
                access_token: 1
            });

        const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
            .reply(400, jsonOut);

        function checkResults() {
            let calls = self.emit.calls;
            expect(calls.count()).toEqual(2);
            expect(calls.argsFor(0)[0]).toEqual('error');
            expect(calls.argsFor(1)[0]).toEqual('end');
            expect(scope1.isDone()).toBeTruthy();
            expect(scope2.isDone()).toBeTruthy();
        }

        action.process.call(self, {
            body: jsonIn
        }, cfg, {})
            .then(checkResults)
            .then(done)
            .catch(done.fail);
    });

    it('should emit error and end events on unsuccessful create request - case: consent problems', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(200, {
                access_token: 1
            });

        const scope2 = nock(microsoftGraphUri).post(microsoftGraphApi)
            .reply(403, jsonOut);

        function checkResults() {
            let calls = self.emit.calls;
            expect(calls.count()).toEqual(2);
            expect(calls.argsFor(0)[0]).toEqual('error');
            expect(calls.argsFor(1)[0]).toEqual('end');
            expect(scope1.isDone()).toBeTruthy();
            expect(scope2.isDone()).toBeTruthy();
        }

        action.process.call(self, {
            body: jsonIn
        }, cfg, {})
            .then(checkResults)
            .then(done)
            .catch(done.fail);
    });

});
