const moment = require('moment');
const chai = require('chai');
const { Logger } = require('@elastic.io/component-commons-library');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/processEventDataHelper');
const data = require('../data/processEventData_test.json');

describe('Outlook Process Event Data', () => {
  it('adds the action config properties to the event post body', (done) => {
    const configInput = data.t1_all_config_values.cfg_in;
    const jsonInput = data.t1_all_config_values.json_in;

    function checkResults(actualOutput) {
      expect(configInput.importance).to.eql(actualOutput.importance);
      expect(configInput.showAs).to.eql(actualOutput.showAs);
      expect(configInput.sensitivity).to.eql(actualOutput.sensitivity);
      expect(configInput.isAllDay).to.eql(actualOutput.isAllDay);
      expect(configInput.timeZone).to.eql(actualOutput.start.timeZone);
      expect(configInput.timeZone).to.eql(actualOutput.end.timeZone);
      expect(configInput.bodyContentType).to.eql(actualOutput.body.contentType);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('does NOT add the isallday config property - when false - to the event post body', (done) => {
    const configInput = data.t2_all_config_values_isAllDay_false.cfg_in;
    const jsonInput = data.t2_all_config_values_isAllDay_false.json_in;

    function checkResults(actualOutput) {
      expect(actualOutput.importance).to.eql(configInput.importance);
      expect(actualOutput.showAs).to.eql(configInput.showAs);
      expect(actualOutput.sensitivity).to.eql(configInput.sensitivity);
      expect(actualOutput.isAllDay).to.eql(undefined);
      expect(actualOutput.start.timeZone).to.eql(configInput.timeZone);
      expect(actualOutput.end.timeZone).to.eql(configInput.timeZone);
      expect(actualOutput.body.contentType).to.eql(configInput.bodyContentType);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('does NOT add any default properties in the event post body', (done) => {
    const configInput = data.t3_no_postbody_config_values.cfg_in;
    const jsonInput = data.t3_no_postbody_config_values.json_in;

    function checkResults(actualOutput) {
      expect(actualOutput.importance).to.eql(undefined);
      expect(actualOutput.showAs).to.eql(undefined);
      expect(actualOutput.sensitivity).to.eql(undefined);
      expect(actualOutput.isAllDay).to.eql(undefined);
      expect(actualOutput.body).to.eql(undefined);
      expect(actualOutput.subject).to.eql(undefined);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('formats start/end time to YYYY-MM-DDTHH:mm:ss for non all day events', (done) => {
    const configInput = data.t6_format_for_non_all_day_events.cfg_in;
    const jsonInput = data.t6_format_for_non_all_day_events.json_in;

    function checkResults(actualOutput) {
      expect(moment(actualOutput.start.dateTime).creationData().format).to.eql('YYYY-MM-DDTHH:mm:ss');
      expect(moment(actualOutput.end.dateTime).creationData().format).to.eql('YYYY-MM-DDTHH:mm:ss');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('does not change ISO start/end time when no utc offset is provided (europe/kiev)', (done) => {
    const configInput = {
      timeZone: 'Europe/Kiev',
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

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2018-09-14T20:00:00');
      expect(actualOutput.end.dateTime).to.eql('2018-09-14T21:00:00');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('does not change start/end time when no utc offset is provided (europe/berlin)', (done) => {
    const configInput = {
      timeZone: 'Europe/Berlin',
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

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2018-09-14T20:00:00');
      expect(actualOutput.end.dateTime).to.eql('2018-09-14T21:00:00');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('converts unix millisec start/end time to iso datetime in cfg timezone (europe/berlin)', (done) => {
    const configInput = {
      timeZone: 'Europe/Berlin',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '1410715640579',
      },
      end: {
        dateTime: '1410715690579',
      },
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2014-09-14T19:27:20');
      expect(actualOutput.end.dateTime).to.eql('2014-09-14T19:28:10');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('converts unix millisec start/end time to iso datetime & cfg timezone (europe/kiev)', (done) => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '1410715640579',
      },
      end: {
        dateTime: '1410715690579',
      },
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2014-09-14T20:27:20');
      expect(actualOutput.end.dateTime).to.eql('2014-09-14T20:28:10');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('changes start/end time when utc offset is provided, to cfg timezone time (europe/berlin)', (done) => {
    const configInput = {
      timeZone: 'Europe/Berlin',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '2016-12-19T18:00:00+02:00',
      },
      end: {
        dateTime: '2016-12-19T19:00:00+02:00',
      },
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2016-12-19T17:00:00');
      expect(actualOutput.end.dateTime).to.eql('2016-12-19T18:00:00');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('changes start/end time when utc offset is provided, to cfg time zone time (europe/kiev)', (done) => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '2016-12-19T18:00:00+02:00',
      },
      end: {
        dateTime: '2016-12-19T19:00:00+02:00',
      },
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2016-12-19T18:00:00');
      expect(actualOutput.end.dateTime).to.eql('2016-12-19T19:00:00');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('adds an extra day to end time for all day events and formats start/end dates as YYYY-MM-DD ', (done) => {
    const configInput = data.t4_add_1_day_for_all_day_events.cfg_in;
    const jsonInput = data.t4_add_1_day_for_all_day_events.json_in;

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2016-12-19');
      expect(actualOutput.end.dateTime).to.eql('2016-12-20');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('processes start/end times values even if the user entered spaces', (done) => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '   2018-09-14T20:00:00        ',
      },
      end: {
        dateTime: '     2018-09-14T21:00:00         ',
      },
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).to.eql('2018-09-14T20:00:00');
      expect(actualOutput.end.dateTime).to.eql('2018-09-14T21:00:00');
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('is rejected when required cfg field calendarId is missing', (done) => {
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

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when required cfg field timeZone is missing', (done) => {
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

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when required input message field start.dateTime is missing', (done) => {
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

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when required input message field end.dateTime is missing', (done) => {
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

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when bodyContentType is provided AND body.content is NOT provided', (done) => {
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

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when an invalid start date is provided', (done) => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: 'this_is_an_invalid_date_string',
      },
      end: {
        dateTime: '2018-09-14T22:00:00',
      },
    };

    const errMessage = 'Invalid date this_is_an_invalid_date_string.';

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when an invalid end date is provided', (done) => {
    const configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '2018-09-14T22:00:00',
      },
      end: {
        dateTime: 'this_is_an_invalid_date_string',
      },
    };

    const errMessage = 'Invalid date this_is_an_invalid_date_string.';

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when a non ISO-8601 format end date is provided', (done) => {
    const configInput = {
      timeZone: 'Europe/Berlin',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: '2016-12-18T20:00',
      },
      end: {
        dateTime: 'Dec 20 2016 09:00:00 GMT+0200',
      },
    };

    const errMessage = 'non ISO-8601 date formats are currently not supported: Dec 20 2016 09:00:00 GMT+0200.';

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when a non ISO-8601 format start date is provided', (done) => {
    const configInput = {
      timeZone: 'Europe/Berlin',
      calendarId: 'gxdfgsdfgID',
    };
    const jsonInput = {
      start: {
        dateTime: 'Dec 19 2016 09:00:00 GMT+0200',
      },
      end: {
        dateTime: '2016-12-18T20:00',
      },
    };

    const errMessage = 'non ISO-8601 date formats are currently not supported: Dec 19 2016 09:00:00 GMT+0200.';

    function checkError(err) {
      expect(err.message).to.eql(errMessage);
    }

    action.processEventData(logger, configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });
});
