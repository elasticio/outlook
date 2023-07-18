const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const _ = require('lodash');
const trigger = require('../../lib/triggers/contacts');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOut = require('../data/getContact_test.out.json');

describe('Outlook Contacts', () => {
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

  it('should emit data without @oauth property - case: http 200', async () => {
    sinon.stub(OutlookClient.prototype, 'getMyLatestContacts').callsFake(() => jsonOut.value);
    await trigger.process.call(self, {}, cfg, {});

    const { callCount, args } = self.emit;
    expect(callCount).to.eql(3);
    expect(args[0][0]).to.eql('data');
    expect(args[0][1].body).to.eql(Object.assign(
      _.omit(jsonOut.value[0], '@odata.etag'),
      {
        calendarId: cfg.calendarId,
      },
    ));
    expect(args[1][0]).to.eql('data');
    expect(args[1][1].body).to.eql(Object.assign(
      _.omit(jsonOut.value[1], '@odata.etag'),
      {
        calendarId: cfg.calendarId,
      },
    ));
  });
});
