'use strict';
const moment = require('moment');

describe('Outlook Process Event Data', function () {

  const action = require('../../lib/processEventDataHelper');
  const data = require('../data/processEventData_test.json');
  const expect = require('chai').expect;

  it('adds the importance/showas/sensitivity/timezone/isallday(when true)/contentbodytype from the cfg param to the event post body', function test(done) {
    let configInput = data.t1_all_config_values.cfg_in;
    let jsonInput = data.t1_all_config_values.json_in;

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect(configInput.importance).to.equal(actualOutput.importance);
      expect(configInput.showAs).to.equal(actualOutput.showAs);
      expect(configInput.sensitivity).to.equal(actualOutput.sensitivity);
      expect(configInput.isAllDay).to.equal(actualOutput.isAllDay);
      expect(configInput.timeZone).to.equal(actualOutput.start.timeZone);
      expect(configInput.timeZone).to.equal(actualOutput.end.timeZone);
      expect(configInput.bodyContentType).to.equal(actualOutput.body.contentType);
    }
  });

  it('does NOT add the isallday (when false) from the cfg param to the event post body', function test(done) {
    let configInput = data.t2_all_config_values_isAllDay_false.cfg_in;
    let jsonInput = data.t2_all_config_values_isAllDay_false.json_in;

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect(configInput.importance).to.equal(actualOutput.importance);
      expect(configInput.showAs).to.equal(actualOutput.showAs);
      expect(configInput.sensitivity).to.equal(actualOutput.sensitivity);
      expect(undefined).to.equal(actualOutput.isAllDay);
      expect(configInput.timeZone).to.equal(actualOutput.start.timeZone);
      expect(configInput.timeZone).to.equal(actualOutput.end.timeZone);
      expect(configInput.bodyContentType).to.equal(actualOutput.body.contentType);
    }
  });

  it('does NOT add any default properties in the event post body when (not required) properties are not defined in CFG', function test(done) {
    let configInput = data.t3_no_postbody_config_values.cfg_in;
    let jsonInput = data.t3_no_postbody_config_values.json_in;

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect(undefined).to.equal(actualOutput.importance);
      expect(undefined).to.equal(actualOutput.showAs);
      expect(undefined).to.equal(actualOutput.sensitivity);
      expect(undefined).to.equal(actualOutput.isAllDay);
      expect(undefined).to.equal(actualOutput.body);
      expect(undefined).to.equal(actualOutput.subject);
    }
  });

  //noinspection Eslint
  it('formats start/end time to YYYY-MM-DD for all day events', function test(done) {
    let configInput = data.t5_format_for_all_day_events.cfg_in;
    let jsonInput = data.t5_format_for_all_day_events.json_in;

    //noinspection Eslint
    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('YYYY-MM-DD').to.equal(moment(actualOutput.start.dateTime).creationData().format);
      expect('YYYY-MM-DD').to.equal(moment(actualOutput.end.dateTime).creationData().format);
    }
  });

  it('formats start/end time to YYYY-MM-DDTHH:mm:ss for non all day events', function test(done) {
    let configInput = data.t6_format_for_non_all_day_events.cfg_in;
    let jsonInput = data.t6_format_for_non_all_day_events.json_in;

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('YYYY-MM-DDTHH:mm:ss').to.equal(moment(actualOutput.start.dateTime).creationData().format);
      expect('YYYY-MM-DDTHH:mm:ss').to.equal(moment(actualOutput.end.dateTime).creationData().format);
    }
  });

  it('does not change start/end time when no utc offset is provided - case: europe/kiev', function test(done) {
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
    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('2018-09-14T20:00:00').to.equal(actualOutput.start.dateTime);
      expect('2018-09-14T21:00:00').to.equal(actualOutput.end.dateTime);
    }
  });

  it('does not change start/end time when no utc offset is provided - case: europe/berlin', function test(done) {
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
    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualOutput) {
      expect('2018-09-14T20:00:00').to.equal(actualOutput.start.dateTime);
      expect('2018-09-14T21:00:00').to.equal(actualOutput.end.dateTime);
    }
  });

  it('converts unix millisec input for start/end time to iso datetime (timeZone as defined in cfg) - case: europe/berlin, isAllDay false', function test(done) {
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
    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('2014-09-14T19:27:20').to.equal(actualOutput.start.dateTime);
      expect('2014-09-14T19:28:10').to.equal(actualOutput.end.dateTime);
    }
});

  it('converts unix millisec input for start/end time to iso datetime (timeZone as defined in cfg) - case: europe/berlin, isAllDay false', function test(done) {
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
    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('2014-09-14T20:27:20').to.equal(actualOutput.start.dateTime);
      expect('2014-09-14T20:28:10').to.equal(actualOutput.end.dateTime);
    }

  });

  it('changes start/end time when utc offset is provided, to the timeZone time as defined in cfg - case: europe/berlin', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('2016-12-19T17:00:00').to.equal(actualOutput.start.dateTime);
      expect('2016-12-19T18:00:00').to.equal(actualOutput.end.dateTime);
    }
  });

  it('changes start/end time when utc offset is provided, to the timeZone time as defined in cfg - case: europe/kiev', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('2016-12-19T18:00:00').to.equal(actualOutput.start.dateTime);
      expect('2016-12-19T19:00:00').to.equal(actualOutput.end.dateTime);
    }
  });

  it('adds an extra day to end time for all day events', function test(done) {
    let configInput = data['t4_add_1_day_for_all_day_events'].cfg_in;
    let jsonInput = data['t4_add_1_day_for_all_day_events'].json_in;
    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done.fail);

    function checkResults(actualOutput) {
      expect('2016-12-19').to.equal(actualOutput.start.dateTime);
      expect('2016-12-20').to.equal(actualOutput.end.dateTime);
    }
  });

  it('processes start/end times values even if the user entered spaces', function test(done) {
    let configInput = {
      timeZone: 'Europe/Kiev',
      calendarId : 'gxdfgsdfgID'
    };
    let jsonInput = {
      start: {
        dateTime: '   2018-09-14T20:00:00        '
      },
      end: {
        dateTime: '     2018-09-14T21:00:00         '
      }
    };
    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .then(done)
      .catch(done);

    function checkResults(actualOutput) {
      expect('2018-09-14T20:00:00').to.equal(actualOutput.start.dateTime);
      expect('2018-09-14T21:00:00').to.equal(actualOutput.end.dateTime);
    }
  });

  it('throws an error when required cfg field calendarId is missing', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualOutput) {
      expect(actualOutput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).to.equal(new Error(errMessage));
    }

   });


  it('throws an error when required cfg field timeZone is missing', function test(done) {
    let configInput = {
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

    let errMessage = 'Time Zone missing! This field is required!';

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualOutput) {
      expect(actualOutput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).to.equal(new Error(errMessage));
    }

  });

  it('throws an error when required input message field start.dateTime is missing', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualOutput) {
      expect(actualOutput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).to.equal(new Error(errMessage));
    }

  });

  it('throws an error when required input message field end.dateTime is missing', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualOutput) {
      expect(actualOutput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).to.equal(new Error(errMessage));
    }

  });

  it('throws an error when bodyContentType is provided in the cfg, but no body.content is provided in the input message', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualOutput) {
      expect(actualOutput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).to.equal(new Error(errMessage));
    }

  });

  it('throws an error when an invalid start date is provided', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualOutput) {
      expect(actualOutput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).to.equal(new Error(errMessage));
    }

  });

  it('throws an error when an invalid end date is provided', function test(done) {
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

    action.processEventData(configInput, jsonInput)
      .then(checkResults)
      .catch(checkError)
      .finally(done);

    function checkResults(actualOutput) {
      expect(actualOutput).toBeUndefined();
    }

    function checkError(err) {
      expect(err).to.equal(new Error(errMessage));
    }

  });


});
