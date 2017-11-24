'use strict';

describe('Outlook Find next available time', function test() {

    const nock = require('nock');
    const action = require('../../lib/actions/findNextAvailableTime');

    const cfg = require('../data/configuration.in.json');
    const jsonIn = require('../data/findNextAvailableTime_test.in.json');
    const jsonOut = require('../data/findNextAvailableTime_test.out.json');

    const refreshTokenUri = 'https://login.microsoftonline.com';
    const refreshTokenApi = '/common/oauth2/v2.0/token';
    const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
    const microsoftGraphApi = '/me/events';


    var self;
    beforeEach(function createSpy() {
        self = jasmine.createSpyObj('self', ['emit']);
    });

    it('should return nextAvailable time on success request - case: http 200', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(200, {
                access_token: 1
            });

        const scope2 = nock(microsoftGraphUri).get(microsoftGraphApi)
            .query({
                $filter: `end/dateTime ge '${jsonIn.time}'`
            })
            .reply(200, jsonOut);

        function checkResults(data) {
            const calls = self.emit.calls;
            expect(calls.count()).toEqual(1);
            expect(calls.argsFor(0)[0]).toEqual('updateKeys');
            expect(calls.argsFor(0)[1]).toEqual({
                oauth: {
                    access_token: 1
                }
            });

            const { time, subject } = jsonIn;

            expect(data.body).toEqual({
                time,
                subject
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

    it('should throw error on unsuccessful refresh token request', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(401, {
                access_token: 1
            });

        action.process.call(self, {
            body: jsonIn
        }, cfg, {})
          .then(() => done.fail(new Error('Error is expected')))
            .catch(err => {
                expect(err).toEqual(jasmine.any(Error));
                expect(err.message).toEqual('Failed to refresh token');
                expect(scope1.isDone()).toBeTruthy();
                done();
            });

    });

});
