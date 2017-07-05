

describe('Outlook Contacts', function test() {

    const nock = require('nock');
    const _ = require('lodash');
    const trigger = require('../../lib/triggers/contacts');

    const cfg = require('../data/configuration.in.json');
    const jsonOut = {
        value: require('../data/getContact_test.out.json')
    };

    const refreshTokenUri = 'https://login.microsoftonline.com';
    const refreshTokenApi = '/common/oauth2/v2.0/token';
    const microsoftGraphUri = 'https://graph.microsoft.com/v1.0';
    const microsoftGraphApi = '/me/contacts';


    var self;
    beforeEach(function createSpy() {
        self = jasmine.createSpyObj('self', ['emit']);
    });

    it('should emit data without @oauth property - case: http 200', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
            .reply(200, {
                access_token: 1
            });

        const scope2 = nock(microsoftGraphUri).get(microsoftGraphApi)
            .query({
                $orderby: 'lastModifiedDateTime asc',
                $top: 900,
                $filter: 'lastModifiedDateTime gt 1970-01-01T00:00:00.000Z'
            })
            .reply(200, jsonOut);

        function checkResults() {
            let calls = self.emit.calls;
            expect(calls.count()).toEqual(4);
            expect(calls.argsFor(0)[0]).toEqual('updateKeys');
            const [firstEvent, firstEmitData] = calls.argsFor(1);
            expect(firstEvent).toBe('data');
            expect(firstEmitData.body).toEqual(Object.assign(
                _.omit(jsonOut.value[0], '@odata.etag'),
                {
                    calendarId: cfg.calendarId
                }
            ));
            const [secondEvent, secondEmitData] = calls.argsFor(2);
            expect(secondEvent).toBe('data');
            expect(secondEmitData.body).toEqual(Object.assign(
                _.omit(jsonOut.value[1], '@odata.etag'),
                {
                    calendarId: cfg.calendarId
                }
            ));
            expect(scope1.isDone()).toBeTruthy();
            expect(scope2.isDone()).toBeTruthy();
        }

        trigger.process.call(self, {}, cfg, {})
            .then(checkResults)
            .then(done)
            .catch(done.fail);
    });

    it('should emit error on unsuccessful refresh token request', done => {
        const scope1 = nock(refreshTokenUri).post(refreshTokenApi)
                                            .reply(401, {
                                                access_token: 1
                                            });

        function checkResults() {
            let calls = self.emit.calls;
            expect(calls.count()).toEqual(1);
            expect(calls.argsFor(0)[0]).toEqual('error');
            expect(calls.argsFor(0)[1]).toEqual(jasmine.any(Error));
            expect(calls.argsFor(0)[1].message).toEqual('Failed to refresh token');
            expect(scope1.isDone()).toBeTruthy();
        }

        trigger.process.call(self, {}, cfg, {})
          .then(checkResults)
          .then(done)
          .catch(done.fail);
    });

});
