Recover from error loading Microsoft Word Document from ShareFile using WOPI.

## Problem

When editing files from ShareFile in Microsoft Word Online, ShareFile passes a time-limited token to Microsoft Word Online via WOPI.
This has information about what ShareFile endpoint to contact to load the file.
However, the token does expire pretty quickly (within a few hours) and is removed from the URI by Microsoft Word Online.
As a result, directly navigating to this URI results in the document failing to connect to ShareFile, displaying the following error:

![Microsoft Word Online WOPI error “Sorry, we ran into a problem.”](https://i.imgur.com/wste0fV.png)

## Solution

When launching a Microsoft Word Online document via WOPI, one navigates to a ShareFile URI that looks like `https://«tenant-name».sharefile.com/e/«document-id»` where `«tenant-name»` is the tenant name (the “subdomain” which ShareFile prompts the user for at [the generalized login page](https://secure.sharefile.com/Authentication/Login)) and `«document-id»` is something that looks like a UUID with its first two characters replaced by `st` (for example, `stf3bc92-0d75-4851-ab9b-6e0a70b7d7a0`).
Navigating to this URI always successfully loads the document, even if the user needs to log into ShareFile first.
So, to recover from the issue, we need to detect that the issue is occurring and then navigate the user to that URI.

The user will encounter the “Sorry, we ran into a problem.” error when directly navigating to the WOPI page.
The URI of the page displaying that error looks something like this:

`https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=https://sfwopionline-usw.sharefile.com/WopiServer/wopi/files/«document-id»&IsLicensedUser=1`

Thus, we can extract `«document-id»`.
However, I do not know of any ShareFile endpoint which is capable of handling a plain `«document-id»`.
It seems that this value is tenant-scoped.
So we need to be able to calculate `«tenant-name»`.
This is impossible.

However, if our script is running, it will modify the WOPI URI once the user arrives at that page to store the tenant name in a GET parameter called `x-sharefile-tenant`.
This way, we can store the value in the user’s browsing history and/or current address bar.
Then, when the user directly navigates to the page via address bar search or reloads/restores the tab, the script can just extract the `«tenant-name»` from the GET parameter.
The URI will look like:

`https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=https://sfwopionline-usw.sharefile.com/WopiServer/wopi/files/«document-id»&IsLicensedUser=1&x-sharefile-tenant=«tenant-name»`

If the error dialog is detected in the document, then the `«document-id»` is extracted from the `WOPISrc` GET parameter and `«tenant-name»` is extracted from the `x-sharefile-tenant` GET parameter.
These are interpolated into `https://«tenant-name».sharefile.com/e/«document-id»` and then the browser is navigated to this URI.

If that value is not present, it may be stil possible to prompt the user for the correct value.
This functionality is not implemented yet.

## Installation

[Install](binki-sharefile-microsoft-word-recovery.user.js?raw=1)
