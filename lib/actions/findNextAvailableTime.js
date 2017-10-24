`use strict`;
const elasticio = require('elasticio-node');
const messages = elasticio.messages;
const MicrosoftGraph = require('msgraph-sdk-javascript');
const ApiClient = require('../apiClient');

module.exports.process = processAction;

async function processAction(msg, cfg) {

    console.log('Refreshing an OAuth Token');

    const instance = new ApiClient(cfg, this);

    //checking if refresh token was successful
    let newAccessToken;
    try {
        newAccessToken = await instance.getRefreshedToken();
    } catch (e) {
        console.log(e);
        throw new Error('Failed to refresh token');
    }

    if (!newAccessToken) {
        return;
    }

    const client = MicrosoftGraph.init({
        defaultVersion: 'v1.0',
        debugLogging: true,
        authProvider: done => {
            done(null, newAccessToken);
        }
    });

    let time = msg.body.time || (new Date()).toISOString();
    let yesterday = new Date('10-23-2017 19:00').toISOString();
    const getEventsResponse = await client
        .api('/me/events')
        .filter(`end/dateTime ge '${time}'`)
        .get();

    let events = getEventsResponse.value.map((e) => ({
        start: new Date(e.start.dateTime),
        end: new Date(e.end.dateTime)
    })).sort((a,b) => a.end - b.end);

    let mimimalTimeBuffer = 15 * 60 * 1000;
    if (events.length === 1) {
        time = events[0].end.toISOString();
    } else {
        let foundPlace = false;
        //from second
        for (let i = 1; i < events.length; i++) {
            if (events[i].start.valueOf()
                - events[i - 1].end.valueOf()
                >= mimimalTimeBuffer) {
                time = events[i - 1].end.toISOString();
                foundPlace = true;
                break;
            }
        }
        if (!foundPlace) {
            time = events[events.length - 1].end.toISOString();
        }
    }
    return messages.newMessageWithBody({
        time,
        subject: msg.body.subject
    });
}
