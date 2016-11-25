'use strict';
const moment = require('moment');

describe('Outlook Process Event Data', function () {

  const action = require('../../lib/processEventDataHelper');
  const data = require('../data/processEventData_test.json');

  it('adds the importance/showas/sensitivity/timezone/isallday(when true)/contentbodytype from the cfg param to the event post body', function () {
    let config_input = data["t1_all_config_values"].cfg_in;
    let json_input = data["t1_all_config_values"].json_in;
    let actual_output = action.processEvent(config_input, json_input);

    expect(config_input.importance).toEqual(actual_output.importance);
    expect(config_input.showAs).toEqual(actual_output.showAs);
    expect(config_input.sensitivity).toEqual(actual_output.sensitivity);
    expect(config_input.isAllDay).toEqual(actual_output.isAllDay);
    expect(config_input.timeZone).toEqual(actual_output.start.timeZone);
    expect(config_input.timeZone).toEqual(actual_output.end.timeZone);
    expect(config_input.bodyContentType).toEqual(actual_output.body.contentType);
  });

  it('does NOT add the isallday (when false) from the cfg param to the event post body', function () {
    let config_input = data["t2_all_config_values_isAllDay_false"].cfg_in;
    let json_input = data["t2_all_config_values_isAllDay_false"].json_in;
    let actual_output = action.processEvent(config_input, json_input);

    expect(config_input.importance).toEqual(actual_output.importance);
    expect(config_input.showAs).toEqual(actual_output.showAs);
    expect(config_input.sensitivity).toEqual(actual_output.sensitivity);
    expect(undefined).toEqual(actual_output.isAllDay);
    expect(config_input.timeZone).toEqual(actual_output.start.timeZone);
    expect(config_input.timeZone).toEqual(actual_output.end.timeZone);
    expect(config_input.bodyContentType).toEqual(actual_output.body.contentType);
  });

  it('does NOT add any default properties in the event post body when properties are not defined in cfg', function () {
    let config_input = data["t3_no_postbody_config_values"].cfg_in;
    let json_input = data["t3_no_postbody_config_values"].json_in;
    let actual_output = action.processEvent(config_input, json_input);

    expect(undefined).toEqual(actual_output.importance);
    expect(undefined).toEqual(actual_output.showAs);
    expect(undefined).toEqual(actual_output.sensitivity);
    expect(undefined).toEqual(actual_output.isAllDay);
    expect(undefined).toEqual(actual_output.start.timeZone);
    expect(undefined).toEqual(actual_output.end.timeZone);
    expect(undefined).toEqual(actual_output.body.contentType);
  });

  it('formats start/end time to YYYY-MM-DD for all day events', function () {
    let config_input = data["t5_format_for_all_day_events"].cfg_in;
    let json_input = data["t5_format_for_all_day_events"].json_in;
    let actual_output = action.processEvent(config_input, json_input);

    expect('YYYY-MM-DD').toEqual(moment(actual_output.start.dateTime).creationData().format);
    expect('YYYY-MM-DD').toEqual(moment(actual_output.end.dateTime).creationData().format);
  });

  it('formats start/end time to YYYY-MM-DDTHH:mm:ss for non all day events', function () {
    let config_input = data["t6_format_for_non_all_day_events"].cfg_in;
    let json_input = data["t6_format_for_non_all_day_events"].json_in;
    let actual_output = action.processEvent(config_input, json_input);

    expect('YYYY-MM-DDTHH:mm:ss').toEqual(moment(actual_output.start.dateTime).creationData().format);
    expect('YYYY-MM-DDTHH:mm:ss').toEqual(moment(actual_output.end.dateTime).creationData().format);;
  });

  it('does not change start/end time when no utc offset is provided - case: europe/kiev', function () {
    let config_input = {
      timeZone: "Europe/Kiev"
    };
    let json_input = {
      start: {
        dateTime: "2018-09-14T20:00:00"
      },
      end: {
        dateTime:  "2018-09-14T21:00:00"
      }
    }
    let actual_output = action.processEvent(config_input, json_input);
    expect('2018-09-14T20:00:00').toEqual(actual_output.start.dateTime);
    expect('2018-09-14T21:00:00').toEqual(actual_output.end.dateTime);
  });

  it('does not change start/end time when no utc offset is provided - case: europe/berlin', function () {
    let config_input = {
      timeZone: "Europe/Berlin"
    };
    let json_input = {
      start: {
        dateTime: "2018-09-14T20:00:00"
      },
      end: {
        dateTime:  "2018-09-14T21:00:00"
      }
    }
    let actual_output = action.processEvent(config_input, json_input);
    expect('2018-09-14T20:00:00').toEqual(actual_output.start.dateTime);
    expect('2018-09-14T21:00:00').toEqual(actual_output.end.dateTime);
  });

  it('converts unix millisec input for start/end time to iso datetime (timeZone as defined in cfg) - case: europe/berlin, isAllDay false', function () {
    let config_input = {
    timeZone: "Europe/Berlin"
  };
  let json_input = {
    start: {
      dateTime: "1410715640579"
    },
    end: {
      dateTime: "1410715690579"
    }
  }
  let actual_output = action.processEvent(config_input, json_input);
  expect('2014-09-14T19:27:20').toEqual(actual_output.start.dateTime);
  expect('2014-09-14T19:28:10').toEqual(actual_output.end.dateTime);

});

  it('converts unix millisec input for start/end time to iso datetime (timeZone as defined in cfg) - case: europe/berlin, isAllDay false', function () {
    let config_input = {
      timeZone: "Europe/Kiev"
    };
    let json_input = {
      start: {
        dateTime: "1410715640579"
      },
      end: {
        dateTime: "1410715690579"
      }
    }
    let actual_output = action.processEvent(config_input, json_input);
    expect('2014-09-14T20:27:20').toEqual(actual_output.start.dateTime);
    expect('2014-09-14T20:28:10').toEqual(actual_output.end.dateTime);

  });

  it('changes start/end time when utc offset is provided, to the timeZone time as defined in cfg - case: europe/berlin', function () {
    let config_input = {
      timeZone: "Europe/Berlin"
    };
    let json_input = {
      start: {
        dateTime: "2016-12-19T18:00:00+02:00"
      },
      end: {
        dateTime: "2016-12-19T19:00:00+02:00"
      }
    }

    let actual_output = action.processEvent(config_input, json_input);
    console.log(actual_output);
    expect("2016-12-19T17:00:00").toEqual(actual_output.start.dateTime);
    expect("2016-12-19T18:00:00").toEqual(actual_output.end.dateTime);
  });

  it('changes start/end time when utc offset is provided, to the timeZone time as defined in cfg - case: europe/kiev', function () {
    let config_input = {
      timeZone: "Europe/Kiev"
    };
    let json_input = {
      start: {
        dateTime: "2016-12-19T18:00:00+02:00"
      },
      end: {
        dateTime: "2016-12-19T19:00:00+02:00"
      }
    }

    let actual_output = action.processEvent(config_input, json_input);
    console.log(actual_output);
    expect("2016-12-19T18:00:00").toEqual(actual_output.start.dateTime);
    expect("2016-12-19T19:00:00").toEqual(actual_output.end.dateTime);
  });

  it('adds an extra day to end time for all day events', function () {
    let config_input = data["t4_add_1_day_for_all_day_events"].cfg_in;
    let json_input = data["t4_add_1_day_for_all_day_events"].json_in;
    let actual_output = action.processEvent(config_input, json_input);
    expect('2016-12-19').toEqual(actual_output.start.dateTime);
    expect('2016-12-20').toEqual(actual_output.end.dateTime);
  });

  it('processes start/end times values even if the user entered spaces', function () {
    let config_input = {
      timeZone: "Europe/Kiev"
    };
    let json_input = {
      start: {
        dateTime: "   2018-09-14T20:00:00        "
      },
      end: {
        dateTime:  "     2018-09-14T21:00:00         "
      }
    }
    let actual_output = action.processEvent(config_input, json_input);
    expect('2018-09-14T20:00:00').toEqual(actual_output.start.dateTime);
    expect('2018-09-14T21:00:00').toEqual(actual_output.end.dateTime);
  });


});
