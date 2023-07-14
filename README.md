[![CircleCI](https://circleci.com/gh/elasticio/outlook.svg?style=svg)](https://circleci.com/gh/elasticio/outlook)

# Outlook component
## Table of Contents

* [General information](#general-information)
   * [Description](#description)
   * [Completeness Matrix](#completeness-matrix)
   * [API version](#api-version)
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
   * [Send Mail](#send-mail)
* [Known Limitations](#known-limitations)

## General information
### Description
[Outlook](https://outlook.live.com/) is a personal information manager web app from Microsoft consisting of webmail, calendaring, contacts, and tasks services.

### Completeness Matrix
![image](https://user-images.githubusercontent.com/16806832/88404425-8a95f400-cdd6-11ea-8712-127d526efbf9.png)

[Completeness Matrix](https://docs.google.com/spreadsheets/d/1fN6keU6GVFGfPSLyNpXilPzsLWJ_leIzAoZZS9BoQ-Y/edit#gid=0)

### API version
The component uses [Microsoft Graph REST API v1.0](https://docs.microsoft.com/en-us/graph/overview?view=graph-rest-1.0).

### Requirements
This component uses [OAuth 2.0 protocol](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols), so you should register your app.
For more details, learn how to register an [app](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).
A Redirect URI for your tenant is: `https://your-tenant.elastic.io/callback/oauth2`, for default EIO tenant just use `https://app.elastic.io/callback/oauth2`.

To create the credentials you would need:
- select existing Auth Client from drop-down list `Choose Auth Client` or create a new one.
  to create a new client you should specify the following fields:

| Field name             | Mandatory | Description                                                                                                              |
|------------------------|-----------|--------------------------------------------------------------------------------------------------------------------------|
| Name                   | true      | your Auth Client's name                                                                                                  |
| Client ID              | true      | your OAuth client key                                                                                                    |
| Client Secret          | true      | your OAuth client secret                                                                                                 |
| Authorization Endpoint | true      | your OAuth authorization endpoint. Use `https://login.microsoftonline.com/common/oauth2/v2.0/authorize`                  |
| Token Endpoint         | true      | your OAuth Token endpoint for refreshing access token. Use `https://login.microsoftonline.com/common/oauth2/v2.0/token`  |

- fill field ``Name Your Credential``
- click on ``Authenticate`` button - if you have not logged in Salesforce before then log in by entering data in the login window that appears
- click on ``Verify`` button for verifying your credentials
- click on ``Save`` button for saving your credentials

This is the list of the scopes that the credentials for the entire component might require. If you want to limit the scope of the credentials, feel free to only select the required scopes (space separated list):
* openid
* offline_access
* User.Read
* Contacts.Read
* Profile
* Calendars.ReadWrite
* Mail.ReadWrite
* Mail.Send

offline_access is required for each credential.
Example of the scopes for the `Send Mail` action: `offline_access Mail.Send`

**Please Note!** To verify be able to verify the credentials you need these scopes: `offline_access User.Read`

### Environment variables
| Name                      |Mandatory|Description|Values|
|---------------------------|---------|-----------|------|
| `LOG_LEVEL`               | false | Controls logger level | `trace`, `debug`, `info`, `warn`, `error` |
| `MAIL_RETRIEVE_MAX_COUNT` | false | Define max count mails could be retrieved per one `Poll for New Mail` trigger execution. Defaults to 1000| 1000 |
| `TOP_LIST_MAIL_FOLDER`    | false | Define the maximum number of folders that can be found for dropdown fields containing a list of Mail Folder. Defaults to 100| 100 |

## Credentials
To create new credentials you need to authorize in Microsoft system using OAuth2 protocol - details are described in [Requirements](#requirements) section.

## Triggers
### Contacts
Triggers to poll all new contacts from Outlook since last polling. Polling is provided by `lastModifiedDateTime` contact's property. 
Per one execution it is possible to poll 900 contacts.
#### Expected output metadata
[/lib/schemas/contacts.out.json](/lib/schemas/contacts.out.json)

### Poll for New Mail
Triggers to poll all new mails from specified folder since last polling. Polling is provided by `lastModifiedDateTime` mail's property. 
Per one execution it is possible to poll 1000 mails by defaults, this can be changed by using environment variable `MAIL_RETRIEVE_MAX_COUNT`.

#### List of Expected Config fields
* **Mail Folder** - Dropdown list with available Outlook mail folders
* **Start Time** - Start datetime of polling. Defaults: `1970-01-01T00:00:00.000Z`
* **Poll Only Unread Mail** - CheckBox, if set, only unread mails will be poll
* **Emit Behavior** -  Options are: default is `Emit Individually` emits each mail in separate message, `Emit All` emits all found mails in one message

#### Expected output metadata
[/lib/schemas/readMail.out.json](/lib/schemas/readMail.out.json)

## Actions

### Check Availability
The action retrieves events for the time specified in `Time` field or for the current time (in case if `Time` field is empty) and returns `true` if no events found, or `false` otherwise.

#### Expected input metadata
[/lib/schemas/checkAvailability.in.json](/lib/schemas/checkAvailability.in.json)
#### Expected output metadata
[/lib/schemas/checkAvailability.out.json](/lib/schemas/checkAvailability.out.json)

### Find Next Available Time
The action retrieves events for the time specified in `Time` field or for the current time (in case if `Time` field is empty).
Returns specified time if no events found, otherwise calculates the new available time based on found event.

#### Expected input metadata
[/lib/schemas/findNextAvailableTime.in.json](/lib/schemas/findNextAvailableTime.in.json)
#### Expected output metadata
[/lib/schemas/findNextAvailableTime.out.json](/lib/schemas/findNextAvailableTime.out.json)

### Create Event
The action creates event in specified calendar with specified options.

#### List of Expected Config fields
* **Calendar** - Dropdown list with available Outlook calendars
* **Time Zone** - Dropdown list with available time zones
* **Importance** - Dropdown list, options are: `Low`, `Normal`, `High`
* **Show As** - Dropdown list, options are: `Free`, `Tentative`, `Busy`, `Out of Office`, `Working Elsewhere`, `Unknown`
* **Sensitivity** - Dropdown list, options are: `Normal`, `Personal`, `Private`, `Confidential`
* **Body Content Type** - Dropdown list, options are: `Text`, `HTML`
* **All Day Event** - CheckBox, if set, all day event will be created

#### Expected input metadata
[/lib/schemas/createEvent.in.json](/lib/schemas/createEvent.in.json)
#### Expected output metadata
[/lib/schemas/createEvent.out.json](/lib/schemas/createEvent.out.json)

### Move Mail
The action moves message with specified id from the original mail folder to a specified destination mail folder or soft-deletes message if the destination folder isn't specified.

#### List of Expected Config fields
* **Original Mail Folder** - Dropdown list with available Outlook mail folders - from where mail should be moved, required field.
* **Destination Mail Folder** - Dropdown list with available Outlook mail folders - where mail should be moved, not required field.
If not specified, the message will be soft-deleted (moved to the folder with property `deleteditems`).

#### Expected input metadata
[/lib/schemas/moveMail.in.json](/lib/schemas/moveMail.in.json)
#### Expected output metadata
Input metadata contains field `Message ID` - what exactly message should be moved.
[/lib/schemas/moveMail.out.json](/lib/schemas/moveMail.out.json)

### Send Mail
The action simply send a message to a recipient(s).

#### Expected input metadata
[/lib/schemas/sendMail.in.json](/lib/schemas/sendMail.in.json)
#### Expected output metadata
In case of a success, output metadata simply repeats the incoming message. I.e. output message schema is exactly the same as for input message.
[/lib/schemas/sendMail.out.json](/lib/schemas/sendMail.in.json)

#### Message Example

```json
{
  "subject": "Hello",
  "toRecipients": [
    {
      "emailAddress": {
        "address": "email@example.com",
        "name": "John"
      }
    }
  ],
  "body": {
    "content": "Hello, I am an email content text",
    "contentType": "text"
  }
}
```

## Known limitations
