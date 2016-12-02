'use strict';
const request = require('request-promise');

function ApiClient(config, requiredScope) {

  const microsoftGraphURI = 'https://graph.microsoft.com/v1.0';
  const refreshTokenURI = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

  function refreshToken() {

    let options = {
      method: 'POST',
      uri: refreshTokenURI,
      json: true,
      form: {
        refresh_token: config.oauth.refresh_token,
        scope: requiredScope,
        grant_type: 'refresh_token',
        client_id: process.env.MSAPP_CLIENT_ID,
        client_secret: process.env.MSAPP_CLIENT_SECRET
      }
    };

    function getAuthorizationHeader(body) {
      return body.access_token;
    }

    return request(options).then(getAuthorizationHeader);
  }

  return {

    post: function postReq(apiCall, requestBody) {

      function setOptions(accessToken) {
        let requestHeaders = {
          Authorization: 'Bearer ' + accessToken
        };
        let options = {
          method: 'POST',
          uri: microsoftGraphURI + apiCall,
          json: true,
          body: requestBody,
          headers: requestHeaders
        };
        return options;

      }

      return refreshToken().then(setOptions).then(request);
    },
    get: function getReq(apiCall) {

      function setOptions(accessToken) {
        let requestHeaders = {
          Authorization: 'Bearer ' + accessToken
        };
        let options = {
          method: 'GET',
          uri: microsoftGraphURI + apiCall,
          json: true,
          headers: requestHeaders
        };
        return options;
      }

      return refreshToken().then(setOptions).then(request);
    }
  };
}

module.exports = ApiClient;
