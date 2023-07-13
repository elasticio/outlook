const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const trigger = require('../../lib/triggers/readMail');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOut = require('../data/readMail_test.out.json');

describe('Outlook Read Mail', () => {
  const folderId = 'folderId';

  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    cfg.folderId = folderId;
    self = {
      emit: sinon.spy(),
      logger,
    };
  });
  afterEach(() => {
    sinon.restore();
  });

  it('should emit data', async () => {
    sinon.stub(OutlookClient.prototype, 'getLatestMessages').callsFake(() => jsonOut.value);

    await trigger.process.call(self, {}, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(3);
    expect(args[0][0]).to.eql('data');
    expect(args[0][1].body).to.eql(jsonOut.value[0]);
    expect(args[1][0]).to.eql('data');
    expect(args[1][1].body).to.eql(jsonOut.value[1]);
    expect(args[2][0]).to.eql('snapshot');
    expect(args[2][1]).to.eql({
      lastModifiedDateTime: '2020-07-20T11:44:53.999Z',
    });
  });

  it('should poll only unread message', async () => {
    sinon.stub(OutlookClient.prototype, 'getLatestMessages').callsFake(() => jsonOut.value);
    cfg.pollOnlyUnreadMail = true;

    await trigger.process.call(self, {}, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(3);
    expect(args[0][0]).to.eql('data');
    expect(args[0][1].body).to.eql(jsonOut.value[0]);
    expect(args[1][0]).to.eql('data');
    expect(args[1][1].body).to.eql(jsonOut.value[1]);
    expect(args[2][0]).to.eql('snapshot');
    expect(args[2][1]).to.eql({
      lastModifiedDateTime: '2020-07-20T11:44:53.999Z',
    });
  });

  it('should poll only unread message and emit all', async () => {
    sinon.stub(OutlookClient.prototype, 'getLatestMessages').callsFake(() => jsonOut.value);
    cfg.pollOnlyUnreadMail = true;
    cfg.emitBehavior = 'emitAll';

    await trigger.process.call(self, {}, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(2);
    expect(args[0][0]).to.eql('data');
    expect(args[0][1].body).to.eql({ results: jsonOut.value });
    expect(args[1][0]).to.eql('snapshot');
    expect(args[1][1]).to.eql({
      lastModifiedDateTime: '2020-07-20T11:44:53.999Z',
    });
  });
});
