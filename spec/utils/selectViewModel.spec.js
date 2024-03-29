const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const { getFolders } = require('../../lib/utils/selectViewModels');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOutMailFolders = require('../data/mailFolders_test.out.json');

describe('Outlook Get Folders', () => {
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

  it('getFolder test', async () => {
    sinon.stub(OutlookClient.prototype, 'getMailFolders').callsFake(() => jsonOutMailFolders.value);
    const result = await getFolders.call(self, cfg);
    expect(result).to.eql({
      1: 'Drafts',
      2: 'Inbox',
    });
  });
});
