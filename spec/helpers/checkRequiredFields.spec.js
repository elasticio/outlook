const chai = require('chai');
const { Logger } = require('@elastic.io/component-commons-library');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/processEventDataHelper');

describe('Outlook Check Required Fields', () => {
  it('throws an error when required cfg field calendarId is missing', () => {
    const configInput = {
      timeZone: 'Europe/Kiev',
    };
    const jsonInput = {
      start: {
        dateTime: '2018-09-14T20:00:00',
      },
      end: {
        dateTime: '2018-09-14T21:00:00',
      },
    };
    const errMessage = 'Calendar ID missing! This field is required!';
    expect(() => {
      action.checkRequiredFields(logger, configInput, jsonInput);
    }).to.throw(errMessage);
  });

  it('throws an error when required cfg field timeZone is missing', () => {
    const configInput = {
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '2018-09-14T20:00:00',
      },
      end: {
        dateTime: '2018-09-14T21:00:00',
      },
    };

    const errMessage = 'Time Zone missing! This field is required!';
    expect(() => {
      action.checkRequiredFields(logger, configInput, jsonInput);
    }).to.throw(errMessage);
  });

  it('throws an error when required input message field start.dateTime is missing', () => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      end: {
        dateTime: '2018-09-14T21:00:00',
      },
    };

    const errMessage = 'Start Time missing! This field is required!';
    expect(() => {
      action.checkRequiredFields(logger, configInput, jsonInput);
    }).to.throw(errMessage);
  });

  it('throws an error when required input message field end.dateTime is missing', () => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '2018-09-14T21:00:00',
      },
    };

    const errMessage = 'End Time missing! This field is required!';
    expect(() => {
      action.checkRequiredFields(logger, configInput, jsonInput);
    }).to.throw(errMessage);
  });

  it('throws an error when bodyContentType is provided AND body.content is NOT provided', () => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
      bodyContentType: 'HTML',
    };
    const jsonInput = {
      start: {
        dateTime: '2018-09-14T21:00:00',
      },
      end: {
        dateTime: '2018-09-14T22:00:00',
      },
    };

    const errMessage = 'Body Type provided, but Body Content is missing!';
    expect(() => {
      action.checkRequiredFields(logger, configInput, jsonInput);
    }).to.throw(errMessage);
  });

  it('throws NO errors when all required fields are provided', () => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
      bodyContentType: 'HTML',
    };
    const jsonInput = {
      body: {
        content: 'test',
      },
      start: {
        dateTime: '2018-09-14T21:00:00',
      },
      end: {
        dateTime: '2018-09-14T22:00:00',
      },
    };

    expect(() => {
      action.checkRequiredFields(logger, configInput, jsonInput);
    }).not.to.throw();
  });
});
