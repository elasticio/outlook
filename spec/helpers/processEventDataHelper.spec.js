'use strict';
const moment = require('moment');

describe('Outlook Process Event Data', function test() {

  const action = require('../../lib/processEventDataHelper');
  const data = require('../data/processEventData_test.json');


  it('adds the action config properties to the event post body', done => {
    let configInput = data.t1_all_config_values.cfg_in;
    let jsonInput = data.t1_all_config_values.json_in;

    function checkResults(actualOutput) {
      expect(configInput.importance).toEqual(actualOutput.importance);
      expect(configInput.showAs).toEqual(actualOutput.showAs);
      expect(configInput.sensitivity).toEqual(actualOutput.sensitivity);
      expect(configInput.isAllDay).toEqual(actualOutput.isAllDay);
      expect(configInput.timeZone).toEqual(actualOutput.start.timeZone);
      expect(configInput.timeZone).toEqual(actualOutput.end.timeZone);
      expect(configInput.bodyContentType).toEqual(actualOutput.body.contentType);
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

  });

  it('does NOT add the isallday config property - when false - to the event post body', done => {
    let configInput = data.t2_all_config_values_isAllDay_false.cfg_in;
    let jsonInput = data.t2_all_config_values_isAllDay_false.json_in;

    function checkResults(actualOutput) {
      expect(actualOutput.importance).toEqual(configInput.importance);
      expect(actualOutput.showAs).toEqual(configInput.showAs);
      expect(actualOutput.sensitivity).toEqual(configInput.sensitivity);
      expect(actualOutput.isAllDay).toEqual(undefined);
      expect(actualOutput.start.timeZone).toEqual(configInput.timeZone);
      expect(actualOutput.end.timeZone).toEqual(configInput.timeZone);
      expect(actualOutput.body.contentType).toEqual(configInput.bodyContentType);
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('does NOT add any default properties in the event post body', done => {
    let configInput = data.t3_no_postbody_config_values.cfg_in;
    let jsonInput = data.t3_no_postbody_config_values.json_in;

    function checkResults(actualOutput) {
      expect(actualOutput.importance).toEqual(undefined);
      expect(actualOutput.showAs).toEqual(undefined);
      expect(actualOutput.sensitivity).toEqual(undefined);
      expect(actualOutput.isAllDay).toEqual(undefined);
      expect(actualOutput.body).toEqual(undefined);
      expect(actualOutput.subject).toEqual(undefined);
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });


  it('formats start/end time to YYYY-MM-DDTHH:mm:ss for non all day events', done => {
    let configInput = data.t6_format_for_non_all_day_events.cfg_in;
    let jsonInput = data.t6_format_for_non_all_day_events.json_in;

    function checkResults(actualOutput) {
      expect(moment(actualOutput.start.dateTime).creationData().format).toEqual('YYYY-MM-DDTHH:mm:ss');
      expect(moment(actualOutput.end.dateTime).creationData().format).toEqual('YYYY-MM-DDTHH:mm:ss');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('does not change ISO start/end time when no utc offset is provided (europe/kiev)', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T20:00:00'
      },
      end: {
        dateTime:  '2018-09-14T21:00:00'
      }
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2018-09-14T20:00:00');
      expect(actualOutput.end.dateTime).toEqual('2018-09-14T21:00:00');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('does not change start/end time when no utc offset is provided (europe/berlin)', done => {
    let configInput = {
      timeZone: 'Europe/Berlin',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T20:00:00'
      },
      end: {
        dateTime: '2018-09-14T21:00:00'
      }
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2018-09-14T20:00:00');
      expect(actualOutput.end.dateTime).toEqual('2018-09-14T21:00:00');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('converts unix millisec start/end time to iso datetime in cfg timezone (europe/berlin)', done => {
    let configInput = {
    timeZone: 'Europe/Berlin',
    calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
     start: {
       dateTime: '1410715640579'
     },
     end: {
       dateTime: '1410715690579'
     }
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2014-09-14T19:27:20');
      expect(actualOutput.end.dateTime).toEqual('2014-09-14T19:28:10');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

  });

  it('converts unix millisec start/end time to iso datetime & cfg timezone (europe/kiev)', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '1410715640579'
      },
      end: {
        dateTime: '1410715690579'
      }
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2014-09-14T20:27:20');
      expect(actualOutput.end.dateTime).toEqual('2014-09-14T20:28:10');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

  });

  it('changes start/end time when utc offset is provided, to cfg timezone time (europe/berlin)', done => {
    let configInput = {
      timeZone: 'Europe/Berlin',
      calendarId: 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2016-12-19T18:00:00+02:00'
      },
      end: {
        dateTime: '2016-12-19T19:00:00+02:00'
      }
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2016-12-19T17:00:00');
      expect(actualOutput.end.dateTime).toEqual('2016-12-19T18:00:00');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

  });

  it('changes start/end time when utc offset is provided, to cfg time zone time (europe/kiev)', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2016-12-19T18:00:00+02:00'
      },
      end: {
        dateTime: '2016-12-19T19:00:00+02:00'
      }
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2016-12-19T18:00:00');
      expect(actualOutput.end.dateTime).toEqual('2016-12-19T19:00:00');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

  });

  it('adds an extra day to end time for all day events and formats start/end dates as YYYY-MM-DD ', done => {
    let configInput = data.t4_add_1_day_for_all_day_events.cfg_in;
    let jsonInput = data.t4_add_1_day_for_all_day_events.json_in;

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2016-12-19');
      expect(actualOutput.end.dateTime).toEqual('2016-12-20');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });

  it('processes start/end times values even if the user entered spaces', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '   2018-09-14T20:00:00        '
      },
      end: {
        dateTime: '     2018-09-14T21:00:00         '
      }
    };

    function checkResults(actualOutput) {
      expect(actualOutput.start.dateTime).toEqual('2018-09-14T20:00:00');
      expect(actualOutput.end.dateTime).toEqual('2018-09-14T21:00:00');
    }

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);
  });


  it('is rejected when required cfg field calendarId is missing', done => {
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

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
   });

  it('is rejected when required cfg field timeZone is missing', done => {
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

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);

  });

  it('is rejected when required input message field start.dateTime is missing', done => {
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

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);

  });

  it('is rejected when required input message field end.dateTime is missing', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T21:00:00'
      }
    };

    let errMessage = 'End Time missing! This field is required!';

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when bodyContentType is provided AND body.content is NOT provided', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID',
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

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);

  });

  it('is rejected when an invalid start date is provided', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: 'this_is_an_invalid_date_string'
      },
      end: {
        dateTime: '2018-09-14T22:00:00'
      }
    };

    let errMessage = 'Invalid date this_is_an_invalid_date_string.';

    function checkError(err) {
     expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);
  });

  it('is rejected when an invalid end date is provided', done => {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId: 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T22:00:00'
      },
      end: {
        dateTime: 'this_is_an_invalid_date_string'
      }
    };

    let errMessage = 'Invalid date this_is_an_invalid_date_string.';

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);

  });

  it('is rejected when a non ISO-8601 format end date is provided', done => {
    let configInput = {
      timeZone: 'Europe/Berlin',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: "2016-12-18T20:00"
      },
      end: {
        dateTime:  "Dec 20 2016 09:00:00 GMT+0200"
      }
    };

    let errMessage = 'non ISO-8601 date formats are currently not supported: Dec 20 2016 09:00:00 GMT+0200.';

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);

  });

  it('is rejected when a non ISO-8601 format start date is provided', done => {
    let configInput = {
      timeZone: 'Europe/Berlin',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: "Dec 19 2016 09:00:00 GMT+0200"
      },
      end: {
        dateTime:  "2016-12-18T20:00"
      }
    };

    let errMessage = 'non ISO-8601 date formats are currently not supported: Dec 19 2016 09:00:00 GMT+0200.';

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

    action.processEventData(configInput, jsonInput)
      .then(() => done.fail(new Error('Error is expected')))
      .catch(checkError)
      .then(done, done.fail);

  });


});
