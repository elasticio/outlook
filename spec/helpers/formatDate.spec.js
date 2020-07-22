const chai = require('chai');

const { expect } = chai;
const action = require('../../lib/processEventDataHelper');

describe('Outlook Format Date', () => {
  // Check if the date FORMAT itself is as expected.
  describe('formats input date string based on given parameter ', () => {
    const timeZone = 'Europe/Berlin';

    it(' - YYYY-MM-DD - ISO Date', () => {
      const inputDate = '2016-12-19T18:00:00';
      const format = 'YYYY-MM-DD';
      const expectedOutput = '2016-12-19';
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });

    it(' - YYYY-MM-DD - millisec input', () => {
      const inputDate = '1410715640579';
      const format = 'YYYY-MM-DD';
      const expectedOutput = '2014-09-14';
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });

    // Format 'YYYY-MM-DDTHH:mm:ss' requires more cases.
    // They are covered in the test below.
  });

  describe('supports ISO date input, converts it based on timezone param, formats it', () => {
    const format = 'YYYY-MM-DDTHH:mm:ss';

    it(' - ISO input date with no UTC offset is not changed (1)', () => {
      const inputDate = '2016-12-19T18:00:00';
      const timeZone = 'Europe/Berlin';
      const expectedOutput = inputDate;
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });

    it(' - ISO input date with no UTC offset is not changed (2)', () => {
      const inputDate = '2016-12-19T18:00:00';
      const timeZone = 'Europe/Kiev';
      const expectedOutput = inputDate;
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });

    it(' - ISO input date with UTC offset is converted (1)', () => {
      const inputDate = '2016-12-19T18:00:00+02:00';
      const timeZone = 'Europe/Berlin';
      const expectedOutput = '2016-12-19T17:00:00';
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });

    it(' - ISO input date with UTC offset is converted (2)', () => {
      const inputDate = '2016-12-19T18:00:00+02:00';
      const timeZone = 'Europe/Kiev';
      const expectedOutput = '2016-12-19T18:00:00';
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });

    it('should parse YYYY-MM-DDTHH:mm:ss.SSSZ', () => {
      const actualOutput = action.formatDate('2017-01-14T10:00:00.000Z', 'Europe/Kiev', format);
      expect(actualOutput).to.eql('2017-01-14T12:00:00');
    });
  });

  // Check if the VALUE of the returned string is as expected.
  describe('supports millisec date input, converts it based on timezone param, formats it', () => {
    const format = 'YYYY-MM-DDTHH:mm:ss';
    const inputDate = '1410715640579';

    it(' - adapts date to given timezone for millisec input', () => {
      const timeZone = 'Europe/Kiev';
      const expectedOutput = '2014-09-14T20:27:20';
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });

    it(' - adapts date to given timezone for millisec input', () => {
      const timeZone = 'Europe/Berlin';
      const expectedOutput = '2014-09-14T19:27:20';
      const actualOutput = action.formatDate(inputDate, timeZone, format);
      expect(actualOutput).to.eql(expectedOutput);
    });
  });

  // Check errors thrown
  describe('throws errors when unsupported or invalid date times are used ', () => {
    const format = 'YYYY-MM-DDTHH:mm:ss';
    const timeZone = 'Europe/Berlin';

    it('- non iso date 1', () => {
      const inputDate = 'Dec 19 2016 18:00:00 GMT+0200';
      expect(() => {
        action.formatDate(inputDate, timeZone, format);
      }).to.throw(`non ISO-8601 date formats are currently not supported: ${inputDate}.`);
    });

    it('- non iso date 2', () => {
      const inputDate = 'October 30, 2014 11:13:00';
      expect(() => {
        action.formatDate(inputDate, timeZone, format);
      }).to.throw(`non ISO-8601 date formats are currently not supported: ${inputDate}.`);
    });

    it('- non iso date 3', () => {
      const inputDate = '2016-12-19 06:00:00 PM';
      expect(() => {
        action.formatDate(inputDate, timeZone, format);
      }).to.throw(`non ISO-8601 date formats are currently not supported: ${inputDate}.`);
    });

    it('- invalid date', () => {
      const inputDate = 'invalid_date';
      expect(() => {
        action.formatDate(inputDate, timeZone, format);
      }).to.throw(`Invalid date ${inputDate}.`);
    });
  });
});
