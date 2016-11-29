'use strict';
const moment = require('moment');

describe('Outlook Process Event Data', function () {

  const action = require('../../lib/processEventDataHelper');
  const data = require('../data/processEventData_test.json');


  it('adds the importance/showas/sensitivity/timezone/isallday(when true)/contentbodytype from the cfg param to the event post body', function (done) {
    let configInput = data['t1_all_config_values'].cfg_in;
    let jsonInput = data['t1_all_config_values'].json_in;

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect(configInput.importance).toEqual(actualInput.importance);
      expect(configInput.showAs).toEqual(actualInput.showAs);
      expect(configInput.sensitivity).toEqual(actualInput.sensitivity);
      expect(configInput.isAllDay).toEqual(actualInput.isAllDay);
      expect(configInput.timeZone).toEqual(actualInput.start.timeZone);
      expect(configInput.timeZone).toEqual(actualInput.end.timeZone);
      expect(configInput.bodyContentType).toEqual(actualInput.body.contentType);
    }
  });

  it('does NOT add the isallday (when false) from the cfg param to the event post body', function (done) {
    let configInput = data['t2_all_config_values_isAllDay_false'].cfg_in;
    let jsonInput = data['t2_all_config_values_isAllDay_false'].json_in;

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect(configInput.importance).toEqual(actualInput.importance);
      expect(configInput.showAs).toEqual(actualInput.showAs);
      expect(configInput.sensitivity).toEqual(actualInput.sensitivity);
      expect(undefined).toEqual(actualInput.isAllDay);
      expect(configInput.timeZone).toEqual(actualInput.start.timeZone);
      expect(configInput.timeZone).toEqual(actualInput.end.timeZone);
      expect(configInput.bodyContentType).toEqual(actualInput.body.contentType);
    }
  });

  it('does NOT add any default properties in the event post body when (not required) properties are not defined in CFG', function (done) {
    let configInput = data['t3_no_postbody_config_values'].cfg_in;
    let jsonInput = data['t3_no_postbody_config_values'].json_in;

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect(undefined).toEqual(actualInput.importance);
      expect(undefined).toEqual(actualInput.showAs);
      expect(undefined).toEqual(actualInput.sensitivity);
      expect(undefined).toEqual(actualInput.isAllDay);
      expect(undefined).toEqual(actualInput.body);
      expect(undefined).toEqual(actualInput.subject);
    }
  });

  it('formats start/end time to YYYY-MM-DD for all day events', function (done) {
    let configInput = data['t5_format_for_all_day_events'].cfg_in;
    let jsonInput = data['t5_format_for_all_day_events'].json_in;

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('YYYY-MM-DD').toEqual(moment(actualInput.start.dateTime).creationData().format);
      expect('YYYY-MM-DD').toEqual(moment(actualInput.end.dateTime).creationData().format);
    }
  });

  it('formats start/end time to YYYY-MM-DDTHH:mm:ss for non all day events', function (done) {
    let configInput = data['t6_format_for_non_all_day_events'].cfg_in;
    let jsonInput = data['t6_format_for_non_all_day_events'].json_in;

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('YYYY-MM-DDTHH:mm:ss').toEqual(moment(actualInput.start.dateTime).creationData().format);
      expect('YYYY-MM-DDTHH:mm:ss').toEqual(moment(actualInput.end.dateTime).creationData().format);
    }
  });

  it('does not change start/end time when no utc offset is provided - case: europe/kiev', function (done) {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T20:00:00'
      },
      end: {
        dateTime:  '2018-09-14T21:00:00'
      }
    };
    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2018-09-14T20:00:00').toEqual(actualInput.start.dateTime);
      expect('2018-09-14T21:00:00').toEqual(actualInput.end.dateTime);
    }
  });

  it('does not change start/end time when no utc offset is provided - case: europe/berlin', function (done) {
    let configInput = {
      timeZone: 'Europe/Berlin',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T20:00:00'
      },
      end: {
        dateTime:  '2018-09-14T21:00:00'
      }
    };
    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2018-09-14T20:00:00').toEqual(actualInput.start.dateTime);
      expect('2018-09-14T21:00:00').toEqual(actualInput.end.dateTime);
    }
  });

  it('converts unix millisec input for start/end time to iso datetime (timeZone as defined in cfg) - case: europe/berlin, isAllDay false', function (done) {
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
    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2014-09-14T19:27:20').toEqual(actualInput.start.dateTime);
      expect('2014-09-14T19:28:10').toEqual(actualInput.end.dateTime);
    }
});

  it('converts unix millisec input for start/end time to iso datetime (timeZone as defined in cfg) - case: europe/berlin, isAllDay false', function (done) {
    let configInput = {
      timeZone: 'Europe/Kiev',
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
    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2014-09-14T20:27:20').toEqual(actualInput.start.dateTime);
      expect('2014-09-14T20:28:10').toEqual(actualInput.end.dateTime);
    }

  });

  it('changes start/end time when utc offset is provided, to the timeZone time as defined in cfg - case: europe/berlin', function (done) {
    let configInput = {
      timeZone: 'Europe/Berlin',
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

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2016-12-19T17:00:00').toEqual(actualInput.start.dateTime);
      expect('2016-12-19T18:00:00').toEqual(actualInput.end.dateTime);
    }
  });

  it('changes start/end time when utc offset is provided, to the timeZone time as defined in cfg - case: europe/kiev', function (done) {
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

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2016-12-19T18:00:00').toEqual(actualInput.start.dateTime);
      expect('2016-12-19T19:00:00').toEqual(actualInput.end.dateTime);
    }
  });

  it('adds an extra day to end time for all day events', function (done) {
    let configInput = data['t4_add_1_day_for_all_day_events'].cfg_in;
    let jsonInput = data['t4_add_1_day_for_all_day_events'].json_in;
    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2016-12-19').toEqual(actualInput.start.dateTime);
      expect('2016-12-20').toEqual(actualInput.end.dateTime);
    }
  });

  it('processes start/end times values even if the user entered spaces', function (done) {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '   2018-09-14T20:00:00        '
      },
      end: {
        dateTime:  '     2018-09-14T21:00:00         '
      }
    };
    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualInput) {
      expect('2018-09-14T20:00:00').toEqual(actualInput.start.dateTime);
      expect('2018-09-14T21:00:00').toEqual(actualInput.end.dateTime);
    }
  });

  it('throws an error when required cfg field calendarId is missing', function (done) {
     let configInput = {
       timeZone: 'Europe/Kiev'
     };
     let jsonInput = {
       start: {
         dateTime: '2018-09-14T20:00:00'
       },
       end: {
         dateTime:  '2018-09-14T21:00:00'
       }
     };

    let errMessage = 'Calendar ID missing! This field is required!';

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualInput) {
      expect(actualInput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

   });


  it('throws an error when required cfg field timeZone is missing', function (done) {
    let configInput = {
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T20:00:00'
      },
      end: {
        dateTime:  '2018-09-14T21:00:00'
      }
    };

    let errMessage = 'Time Zone missing! This field is required!';

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualInput) {
      expect(actualInput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

  });

  it('throws an error when required input message field start.dateTime is missing', function (done) {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      end: {
        dateTime: '2018-09-14T21:00:00'
      }
    };

    let errMessage = 'Start Time missing! This field is required!';

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualInput) {
      expect(actualInput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

  });

  it('throws an error when required input message field end.dateTime is missing', function (done) {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T21:00:00'
      }
    };

    let errMessage = 'Ent Time missing! This field is required!';

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualInput) {
      expect(actualInput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

  });

  it('throws an error when bodyContentType is provided in the cfg, but no body.content is provided in the input message', function (done) {
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

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualInput) {
      expect(actualInput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

  });

  it('throws an error when an invalid start date is provided', function (done) {
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

    let errMessage = 'Invalid date this_is_an_invalid_date_string';

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualInput) {
      expect(actualInput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

  });

  it('throws an error when an invalid end date is provided', function (done) {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '2018-09-14T22:00:00'
      },
      end: {
        dateTime: 'this_is_an_invalid_date_string'
      }
    };

    let errMessage = 'Invalid date this_is_an_invalid_date_string';

    action.processEvent(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualInput) {
      expect(actualInput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).toEqual(new Error(errMessage));
    }

  });


});
