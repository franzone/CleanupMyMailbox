# CleanupMyMailbox
CleanupMyMailbox is a command-line utility that deletes emails of a certain age from a specific mailbox within Microsoft Exchange Server.

## Background
I work within a corporate environment that manages its own Microsoft Exchange Server. I am also on a development team and receive thousands
of automated email messages per day from various services that run across multiple servers. I started out creating rules for these emails
in MS Outlook that sorted them into folders based on the service that generated them. I would then go in once a day and clean out the folders.
This is a bit tedious. Additionally, I want to keep emails that are within the last day in case something goes wrong with that service and
I need to reference an email that was generated. Even more tedious.

Next I thought to myself, _"Why don't I use the built-in Outlook Archive method?"_ That would be a great option, but my company manages our
archive policies. That means I cannot modify them.

Finally, having all other avenues closed to me, I decided to write my own utility to perform this task. I am, after all, a Software Engineer.

## Download
You can click on the __Release__ tab above or navigate to https://github.com/franzone/CleanupMyMailbox/releases.

## Usage
### To print help
```
.\CleanupMyMailbox.exe --help
```
This will print the following:
```
********************************************************************************
Cleanup My Mailbox 1.0.0
Copyright (c) 2020 Jonathan Franzone
********************************************************************************

  -v, --verbose               Set output to verbose messages.
  -e, --email                 Required. Email address
  -m, --mailbox               Path to the mailbox that needs to be cleaned. Use
                              a forward slash path separator. For example:
                              INBOX/Automated/Services
  -d, --defaultcredentials    Use default credentials. Use this option for
                              hosted Exchange instances where the user is logged
                              into the same domain.
  -u, --username              Username for authenticating to the Exchange
                              server. Leave blank if the same as your email
                              address.
  -s, --password              Password for authenticating to the Exchange
                              server. If required and left blank the program
                              will prompt for a password.
  -a, --age                   Required. The age of emails after which they
                              should be deleted. This should be a string that
                              the .NET TimeSpan.Parse() method will recognize.
                              For example: 1 = 1 day, 6:30 = 6 hours and 30
                              minutes
  -x, --deletemode            Sets the Delete mode to Hard Delete. Emails will
                              not go into Deleted Items folders
  -n, --noprompt              Do not prompt before deleting emails
  --help                      Display this help screen.
  --version                   Display version information.
```

### Within On-premises Exchange Environment
If you are targeting an on-premises Exchange server and your client/workstation is domain joined, then you don't need to enter credentials.
```
.\CleanupMyMailbox.exe --email my.email@email.org --mailbox "Sorted/Service Emails" --age 1 --defaultcredentials
```
This will delete emails in the folder `INBOX\Sorted\Service Emails` that are at least 1 day old for the email `my.email@email.org`.

### To Connect Remotely
If you are connecting to a remote instance of Exchange (i.e. Office365), then you'll need to enter credentials.
```
.\CleanupMyMailbox.exe --email my.email@email.org --mailbox "Sorted/Service Emails" --age 1 --password MyS3cret
```
You may use the `--username` parameter to specify your login username. If this is not provided then the utility will use your email address
(as provided by `--email`). You may use the `--password` parameter to provide your email password on the command line. If you do not provide
the `--password` then the utility will prompt (if it is required).

## Limitations
* Currently limited to processing 1000 emails at a time. This is because of a paging feature in EWS. I haven't bothered to implement
  any type of paging.
* It is possible that the list of emails found in the search could change while processing. This program makes no attempt to detect that
  condition or to correctly handle it.

## Term & Conditions of Usage
By using this software you assume all responsibility for anything and everything that could ever go wrong... ever. So don't sue me if you break something.

## References
* [How to Delete Old Emails in MS Exchange Using EWS](http://tech.franzone.blog/2020/02/05/how-to-delete-old-emails-in-ms-exchange-using-ews/)
* [Get started with EWS Managed API client applications](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
* [TimeSpan.Parse Method](https://docs.microsoft.com/en-us/dotnet/api/system.timespan.parse?view=netframework-4.6.2) - See his page for
  the string format for various timespans.
