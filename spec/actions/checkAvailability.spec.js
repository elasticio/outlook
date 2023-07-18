const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();

const configuration = require('../data/configuration.new.in.json');
const jsonIn = require('../data/checkAvailability_test.in.json');
const jsonOut = require('../data/checkAvailability_test.out.json');
const jsonWithDataOut = require('../data/checkAvailability_withDataTest.out.json');

const cfgString = JSON.stringify(configuration);

const action = require('../../lib/actions/checkAvailability');
const { OutlookClient } = require('../../lib/OutlookClient');

describe('Outlook Check Availability', () => {
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

  it('should return available=true on success request - case: there is no events', async () => {
    sinon.stub(OutlookClient.prototype, 'getMyLatestEvents').callsFake(() => jsonOut);
    await action.process.call(self, { body: jsonIn }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(1);
    expect(args[0][0]).to.eql('data');
    expect(args[0][1].body).to.eql({
      available: true,
    });
  });

  it('should return available=false on success request - case: there are events', async () => {
    sinon.stub(OutlookClient.prototype, 'getMyLatestEvents').callsFake(() => jsonWithDataOut);
    await action.process.call(self, { body: jsonIn }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(1);
    expect(args[0][0]).to.eql('data');
    expect(args[0][1].body).to.eql({
      available: false,
    });
  });
});
