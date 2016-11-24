'use strict';
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

  it('adds an extra day to end time for all day events', function () {
    let config_input = data["t4_add_1_day_for_all_day_events"].cfg_in;
    let json_input = data["t4_add_1_day_for_all_day_events"].json_in;
    let actual_output = action.processEvent(config_input, json_input);
    expect('2016-12-19').toEqual(actual_output.start.dateTime);
    expect('2016-12-20').toEqual(actual_output.end.dateTime);
  });

  it('formats start and end time to YYYY-MM-DD for all day events', function () {
    let config_input = data["t5_format_for_all_day_events"].cfg_in;
    let json_input = data["t5_format_for_all_day_events"].json_in;
    let actual_output = action.processEvent(config_input, json_input);

    expect('2016-12-19').toEqual(actual_output.start.dateTime);
    expect('2016-12-20').toEqual(actual_output.end.dateTime);
  });

  it('formats start and end time to YYYY-MM-DDTHH:mm:ss for non all day events', function () {
    let config_input = data["t6_format_for_non_all_day_events"].cfg_in;
    let json_input = data["t6_format_for_non_all_day_events"].json_in;
    let actual_output = action.processEvent(config_input, json_input);

    expect('2016-12-19T18:00:00').toEqual(actual_output.start.dateTime);
    expect('2016-12-20T21:00:00').toEqual(actual_output.end.dateTime);
  });

  it('formats unix millisec format start and end time to YYYY-MM-DDTHH:mm:ss for non all day events', function () {
    let config_input = {};
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

  



});
