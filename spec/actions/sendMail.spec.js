const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/sendMail');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonIn = require('../data/sendMail_test.in.json');

describe('Outlook Send Mail', () => {
  const msg = {
    body: {
      jsonIn,
    },
  };

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

  it('should send message', async () => {
    sinon.stub(OutlookClient.prototype, 'sendMail').callsFake(() => {});
    const result = await action.process.call(self, msg, cfg, {});
    expect(result.body).to.eql(msg.body);
  });
});
