# CleanupMyMailbox
CleanupMyMailbox is a command-line utility that deletes emails of a certain age from a specific mailbox within Microsoft Exchange Server.

## Background
I work within a corporate environment that manages its own Microsoft Exchange Server. I am also on a development team and receive thousands of automated email messages per day from various services that run across multiple servers. I started out creating rules for these emails in MS Outlook that sorted them into folders based on the service that generated them. I would then go in once a day and clean out the folders. This is a bit tedious. Additionally, I want to keep emails that are within the last day in case something goes wrong with that service and I need to reference an email that was generated. Even more tedious.

## Usage
### To print help
```
.\CleanupMyMailbox.exe --help
```

## References
* https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications
