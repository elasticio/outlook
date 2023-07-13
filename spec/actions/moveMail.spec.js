const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/moveMail');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');

const cfgString = JSON.stringify(configuration);
const jsonOut = require('../data/moveMail_test.out.json');
const folder = require('../data/moveMail_folder_test.in.json');

describe('Outlook Move Mail', () => {
  const originalMailFolders = 'originalId';
  const destinationFolder = 'destinationId';
  const messageId = 'messageId';
  const msg = {
    body: {
      messageId,
    },
  };

  let self;
  let cfg;
  beforeEach(() => {
    cfg = JSON.parse(cfgString);
    cfg.originalMailFolders = originalMailFolders;
    self = {
      emit: sinon.spy(),
      logger,
    };
  });
  afterEach(() => {
    sinon.restore();
  });

  it('should move message to specified folder case: destinationFolder is specified', async () => {
    sinon.stub(OutlookClient.prototype, 'getMailFolderById').callsFake(() => folder);
    sinon.stub(OutlookClient.prototype, 'moveMessage').callsFake(() => jsonOut);
    cfg.destinationFolder = destinationFolder;
    const result = await action.process.call(self, msg, cfg, {});
    const expectedResult = jsonOut;
    expectedResult.currentFolder = {
      id: destinationFolder,
    };
    expect(result.body).to.eql(expectedResult);
  });

  it('should move message to Deleted Items folder', async () => {
    sinon.stub(OutlookClient.prototype, 'getMailFolderById').callsFake(() => folder);
    folder.id = 'deletedItemsFolderId';
    sinon.stub(OutlookClient.prototype, 'getDeletedItemsFolder').callsFake(() => folder);
    sinon.stub(OutlookClient.prototype, 'moveMessage').callsFake(() => jsonOut);

    const result = await action.process.call(self, msg, cfg, {});
    const expectedResult = jsonOut;
    expectedResult.currentFolder = {
      id: 'deletedItemsFolderId',
    };
    expect(result.body).to.eql(expectedResult);
  });
});
