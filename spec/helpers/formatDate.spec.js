'use strict';

describe('Outlook Format Date', function test() {

    const action = require('../../lib/processEventDataHelper');

  //Check if the date FORMAT itself is as expected.
    describe('formats input date string based on given parameter ', function test() {

        let timeZone = 'Europe/Berlin';

        it(' - YYYY-MM-DD - ISO Date', function test() {
            let inputDate = '2016-12-19T18:00:00';
            let format = 'YYYY-MM-DD';
            let expectedOutput = '2016-12-19';
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

        it(' - YYYY-MM-DD - millisec input', function test() {
            let inputDate = '1410715640579';
            let format = 'YYYY-MM-DD';
            let expectedOutput = '2014-09-14';
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

    //Format 'YYYY-MM-DDTHH:mm:ss' requires more cases.
    //They are covered in the test below.
    });


    describe('supports ISO date input, converts it based on timezone param, formats it', function test() {

        let format = 'YYYY-MM-DDTHH:mm:ss';

        it(' - ISO input date with no UTC offset is not changed (1)', function test() {
            let inputDate = '2016-12-19T18:00:00';
            let timeZone = 'Europe/Berlin';
            let expectedOutput = inputDate;
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

        it(' - ISO input date with no UTC offset is not changed (2)', function test() {
            let inputDate = '2016-12-19T18:00:00';
            let timeZone = 'Europe/Kiev';
            let expectedOutput = inputDate;
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

        it(' - ISO input date with UTC offset is converted (1)', function test() {
            let inputDate = '2016-12-19T18:00:00+02:00';
            let timeZone = 'Europe/Berlin';
            let expectedOutput = '2016-12-19T17:00:00';
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

        it(' - ISO input date with UTC offset is converted (2)', function test() {
            let inputDate = '2016-12-19T18:00:00+02:00';
            let timeZone = 'Europe/Kiev';
            let expectedOutput = '2016-12-19T18:00:00';
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

        it('should parse YYYY-MM-DDTHH:mm:ss.SSSZ', () => {
            const actualOutput = action.formatDate('2017-01-14T10:00:00.000Z', 'Europe/Kiev', format);
            expect(actualOutput).toEqual('2017-01-14T12:00:00');
        });

    });

  //Check if the VALUE of the returned string is as expected.
    describe('supports millisec date input, converts it based on timezone param, formats it', function test() {

        let format = 'YYYY-MM-DDTHH:mm:ss';
        let inputDate = '1410715640579';

        it(' - adapts date to given timezone for millisec input', function test() {
            let timeZone = 'Europe/Kiev';
            let expectedOutput = '2014-09-14T20:27:20';
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

        it(' - adapts date to given timezone for millisec input', function test() {
            let timeZone = 'Europe/Berlin';
            let expectedOutput = '2014-09-14T19:27:20';
            let actualOutput = action.formatDate(inputDate, timeZone, format);
            expect(actualOutput).toEqual(expectedOutput);
        });

    });

  //Check errors thrown
    describe('throws errors when unsupported or invalid date times are used ', function test() {
        let format = 'YYYY-MM-DDTHH:mm:ss';
        let timeZone = 'Europe/Berlin';

        it('- non iso date 1', function test() {
            let inputDate = 'Dec 19 2016 18:00:00 GMT+0200';
            expect(function check() {
                action.formatDate(inputDate, timeZone, format);
            }).toThrow(new Error(`non ISO-8601 date formats are currently not supported: ${inputDate}.`));
        });

        it('- non iso date 2', function test() {
            let inputDate = 'October 30, 2014 11:13:00';
            expect(function check() {
                action.formatDate(inputDate, timeZone, format);
            }).toThrow(new Error(`non ISO-8601 date formats are currently not supported: ${inputDate}.`));
        });

        it('- non iso date 3', function test() {
            let inputDate = '2016-12-19 06:00:00 PM';
            expect(function check() {
                action.formatDate(inputDate, timeZone, format);
            }).toThrow(new Error(`non ISO-8601 date formats are currently not supported: ${inputDate}.`));
        });

        it('- invalid date', function test() {
            let inputDate = 'invalid_date';
            expect(function check() {
                action.formatDate(inputDate, timeZone, format);
            }).toThrow(new Error(`Invalid date ${inputDate}.`));
        });

    });

});
