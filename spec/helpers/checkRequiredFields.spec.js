'use strict';

describe('Outlook Check Required Fields', function test() {

    const action = require('../../lib/processEventDataHelper');

    it('throws an error when required cfg field calendarId is missing', function test() {
        let configInput = {
            timeZone: 'Europe/Kiev'
        };
        let jsonInput = {
            start: {
                dateTime: '2018-09-14T20:00:00'
            },
            end: {
                dateTime: '2018-09-14T21:00:00'
            }
        };
        let errMessage = 'Calendar ID missing! This field is required!';
        expect(function check() {
            action.checkRequiredFields(configInput, jsonInput);
        }).toThrow(new Error(errMessage));
    });

    it('throws an error when required cfg field timeZone is missing', function test() {
        let configInput = {
            calendarId: 'gxdfgsdfgID'
        };
        let jsonInput = {
            start: {
                dateTime: '2018-09-14T20:00:00'
            },
            end: {
                dateTime: '2018-09-14T21:00:00'
            }
        };

        let errMessage = 'Time Zone missing! This field is required!';
        expect(function check() {
            action.checkRequiredFields(configInput, jsonInput);
        }).toThrow(new Error(errMessage));

    });

    it('throws an error when required input message field start.dateTime is missing', function test() {
        let configInput = {
            timeZone: 'Europe/Kiev',
            calendarId: 'gxdfgsdfgID'
        };
        let jsonInput = {
            end: {
                dateTime: '2018-09-14T21:00:00'
            }
        };

        let errMessage = 'Start Time missing! This field is required!';
        expect(function check() {
            action.checkRequiredFields(configInput, jsonInput);
        }).toThrow(new Error(errMessage));
    });

    it('throws an error when required input message field end.dateTime is missing', function test() {
        let configInput = {
            timeZone: 'Europe/Kiev',
            calendarId: 'gxdfgsdfgID'
        };
        let jsonInput = {
            start: {
                dateTime: '2018-09-14T21:00:00'
            }
        };

        let errMessage = 'End Time missing! This field is required!';
        expect(function check() {
            action.checkRequiredFields(configInput, jsonInput);
        }).toThrow(new Error(errMessage));

    });

    it('throws an error when bodyContentType is provided AND body.content is NOT provided', function test() {
        let configInput = {
            timeZone: 'Europe/Kiev',
            calendarId: 'gxdfgsdfgID',
            bodyContentType: 'HTML'
        };
        let jsonInput = {
            start: {
                dateTime: '2018-09-14T21:00:00'
            },
            end: {
                dateTime: '2018-09-14T22:00:00'
            }
        };

        let errMessage = 'Body Type provided, but Body Content is missing!';
        expect(function check() {
            action.checkRequiredFields(configInput, jsonInput);
        }).toThrow(new Error(errMessage));

    });

    it('throws NO errors when all required fields are provided', function test() {
        let configInput = {
            timeZone: 'Europe/Kiev',
            calendarId: 'gxdfgsdfgID',
            bodyContentType: 'HTML'
        };
        let jsonInput = {
            body: {
                content: 'test'
            },
            start: {
                dateTime: '2018-09-14T21:00:00'
            },
            end: {
                dateTime: '2018-09-14T22:00:00'
            }
        };

        expect(function check() {
            action.checkRequiredFields(configInput, jsonInput);
        }).not.toThrow();

    });

});
