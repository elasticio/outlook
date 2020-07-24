[![CircleCI](https://circleci.com/gh/elasticio/outlook.svg?style=svg)](https://circleci.com/gh/elasticio/outlook)

# Outlook component
## Table of Contents

* [General information](#general-information)
   * [Description](#description)
   * [Completeness Matrix](#completeness-matrix)
   * [API version / SDK version](#api-version--sdk-versio)
   * [Requirements](#requirements)
   * [Environment variables](#environment-variables)
* [Credentials](#credentials)
* [Triggers](#triggers)
   * [Contacts](#contacts)
   * [Poll for New Mail](#poll-for-new-mail)
* [Actions](#actions)
   * [Check Availability](#check-availability)
   * [Find Next Available Time](#find-next-available-time)
   * [Create Event](#create-event)
   * [Move Mail](#move-mail)
* [Known Limitations](#known-limitations)


## General information
### Description
[Outlook](https://outlook.live.com/) is a personal information manager web app from Microsoft consisting of webmail, calendaring, contacts, and tasks services.

### Completeness Matrix
![image](https://user-images.githubusercontent.com/16806832/88404425-8a95f400-cdd6-11ea-8712-127d526efbf9.png)

[Completeness Matrix](https://docs.google.com/spreadsheets/d/1fN6keU6GVFGfPSLyNpXilPzsLWJ_leIzAoZZS9BoQ-Y/edit#gid=0)

### API version / SDK version
It is used [Microsoft Graph REST API v1.0](https://docs.microsoft.com/en-us/graph/overview?view=graph-rest-1.0).

### Requirements
This component uses [OAuth 2.0 protocol](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols), so you should register your app.
For more details, learn how to register an [app](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).
A Redirect URI for your tenant is: `https://your-tenant.elastic.io/callback/oauth2`, for default EIO tenant just use `https://app.elastic.io/callback/oauth2`.
Client ID and Secret (that you get after app registration) need to be configured in the environment variables ```OAUTH_CLIENT_ID``` and ```OAUTH_CLIENT_SECRET```

### Environment variables
Name|Mandatory|Description|Values|
|----|---------|-----------|------|
|`OAUTH_CLIENT_ID`| true | Microsoft Graph Application OAuth2 Client ID | Can be found in your application page on [https://portal.azure.com](https://portal.azure.com) |
|`OAUTH_CLIENT_SECRET`| true | Microsoft Graph Application OAuth2 Client Secret | Can be found in your application page on [https://portal.azure.com](https://portal.azure.com) |
|`LOG_LEVEL`| false | Controls logger level | `trace`, `debug`, `info`, `warn`, `error` |
|`MAIL_RETRIEVE_MAX_COUNT`| false | Define max count mails could be retrieved per one `Poll for New Mail` trigger execution. Default to 1000| 1000 |

## Credentials
To create new credentials you need to authorize in Microsoft system using OAuth2 protocol - details are described in [Requirements](#requirements) section.

## Triggers
### Contacts
#### List of Expected Config fields
#### Expected output metadata
[/lib/schemas/contacts.out.json](/lib/schemas/contacts.out.json)

### Poll for New Mail
#### List of Expected Config fields
#### Expected output metadata
[/lib/schemas/readMail.out.json](/lib/schemas/readMail.out.json)

## Actions

### Check Availability
#### List of Expected Config fields
#### Expected input metadata
[/lib/schemas/checkAvailability.in.json](/lib/schemas/checkAvailability.in.json)
#### Expected output metadata
[/lib/schemas/checkAvailability.out.json](/lib/schemas/checkAvailability.out.json)
#### Sample pseudo-code (optional)
#### Known limitations for the particular trigger/action / Planned future stages
#### Links to trigger/action specific documentation

### Find Next Available Time
#### List of Expected Config fields
#### Expected input metadata
[/lib/schemas/findNextAvailableTime.in.json](/lib/schemas/findNextAvailableTime.in.json)
#### Expected output metadata
[/lib/schemas/findNextAvailableTime.out.json](/lib/schemas/findNextAvailableTime.out.json)
#### Sample pseudo-code (optional)
#### Known limitations for the particular trigger/action / Planned future stages
#### Links to trigger/action specific documentation

### Create Event
#### List of Expected Config fields
#### Expected input metadata
[/lib/schemas/createEvent.in.json](/lib/schemas/createEvent.in.json)
#### Expected output metadata
[/lib/schemas/createEvent.out.json](/lib/schemas/createEvent.out.json)
#### Sample pseudo-code (optional)
#### Known limitations for the particular trigger/action / Planned future stages
#### Links to trigger/action specific documentation

### Move Mail
#### List of Expected Config fields
#### Expected input metadata
[/lib/schemas/moveMail.in.json](/lib/schemas/moveMail.in.json)
#### Expected output metadata
[/lib/schemas/moveMail.out.json](/lib/schemas/moveMail.out.json)
#### Sample pseudo-code (optional)
#### Known limitations for the particular trigger/action / Planned future stages
#### Links to trigger/action specific documentation

## Known limitations (common for the component)
