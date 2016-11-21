'use strict';
const Helper = require('./helper');

module.exports = function baseDynamicSelectModelFunction(apiCall, processData, cfg, requiredScopes, cb) {

  const instance = new Helper(cfg, requiredScopes);

  function onSuccess(data) {
    cb(null, processData(data.value));
  }

  function onError(err) {
    console.log('Oops! Error occurred');
    console.log(err.message);
    cb(err);
  }

  instance.get(apiCall).then(onSuccess).fail(onError).done();

};

