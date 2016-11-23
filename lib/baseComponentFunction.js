'use strict';
const Helper = require('./helper');

module.exports = function baseComponentFunction(componentObject, apiCall, postBody, emitData, cfg, requiredScopes) {

  const self = componentObject;
  const instance = new Helper(cfg, requiredScopes);

  function emitError(e) {
    console.log('Oops! Error occurred');
    self.emit('error', e);
  }

  function emitEnd() {
    console.log('Finished execution!');
    self.emit('end');
  }


  //TODO: Separate GET/POST/... etc - the method used must be visible in the action file itself
  if (postBody !== null) {
   instance.post(apiCall, postBody)
            .then(emitData)
            .fail(emitError)
            .done(emitEnd);
  } else {
   instance.get(apiCall)
           .then(emitData)
           .fail(emitError)
           .done(emitEnd);
  }

};

