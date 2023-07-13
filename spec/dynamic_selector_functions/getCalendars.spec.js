const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const { getCalendars } = require('../../lib/actions/createEvent');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');
const jsonOut = require('../data/getCalendars_test.out.json');

const cfgString = JSON.stringify(configuration);
describe('Outlook Get Calendars', () => {
  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    self = {
      emit: sinon.spy(),
      logger,
    };
  });

  it('should return calendar info on success get request', async () => {
    sinon.stub(OutlookClient.prototype, 'getMyCalendars').callsFake(() => jsonOut.value);
    const result = await getCalendars.call(self, cfg);
    expect(result).to.eql({
      'AAMkAGI2TGuLAAA=': 'Calendar',
    });
  });
});
