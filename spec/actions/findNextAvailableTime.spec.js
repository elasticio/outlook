const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/findNextAvailableTime');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');
const jsonIn = require('../data/findNextAvailableTime_test.in.json');
const jsonOut = require('../data/findNextAvailableTime_test.out.json');

const cfgString = JSON.stringify(configuration);

describe('Outlook Find next available time', () => {
  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    self = {
      emit: sinon.spy(),
      logger,
    };
  });
  afterEach(() => {
    sinon.restore();
  });

  it('should return nextAvailable time on success request - case: http 200', async () => {
    sinon.stub(OutlookClient.prototype, 'getMyLatestEvents').callsFake(() => jsonOut);
    const result = await action.process.call(self, { body: jsonIn }, cfg, {});
    expect(result.body).to.eql({
      time: jsonIn.time,
      subject: jsonIn.subject,
    });
  });
});
