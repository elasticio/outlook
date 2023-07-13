const { Logger } = require('@elastic.io/component-commons-library');
const chai = require('chai');
const sinon = require('sinon');
require('../commons');

const { expect } = chai;
const logger = Logger.getLogger();
const action = require('../../lib/actions/createEvent');
const { OutlookClient } = require('../../lib/OutlookClient');

const configuration = require('../data/configuration.new.in.json');
const jsonIn = require('../data/createEvent_test.in.json');
const jsonOut = require('../data/createEvent_test.out.json');

const cfgString = JSON.stringify(configuration);

describe('Outlook Create Event', () => {
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

  it('should emit (data and end events on success create request - case: event created', async () => {
    sinon.stub(OutlookClient.prototype, 'createEvent').callsFake(() => jsonOut);

    await action.process.call(self, { body: jsonIn }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.be.eql(1);
    expect(args[0][0]).to.be.eql('data');
    expect(args[0][1].body).to.be.eql({
      id: 'testid12345',
      subject: 'Unit Test - Simple Event',
      body: {
        contentType: 'HTML',
        content: 'This is a test.',
      },
      start: {
        dateTime: '2017-05-19T18:00:00',
        timeZone: 'Central European Standard Time',
      },
      end: {
        dateTime: '2017-05-20T19:00:00',
        timeZone: 'Central European Standard Time',
      },
      calendarId: 'AAMkAGYyNmJlYjBmLTgwOWYtNGU0Mi04NW',
    });
  });

  it('should emit error and end events on unsuccessful create request - case: bad request', async () => {
    sinon.stub(OutlookClient.prototype, 'createEvent').throws(new Error('Bad Request'));

    await action.process.call(self, {
      body: jsonIn,
    }, cfg, {});
    const { callCount, args } = self.emit;
    expect(callCount).to.eql(1);
    expect(args[0][0]).to.eql('error');
  });
});
