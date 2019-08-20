# MAPIRecipientsTableSample
A code sample that illustrates how to access and update the recipients collection on Outlook items. The sample prompts for the profile to logon to. 

# Warning
Please do not run the exe on production mailboxes as it will result in chages to the recipients table. 

# Sample output

```s
Opening item with subject "Test appointment"
        The selected mesasge has no recipients.

Opening item with subject "test 1"
        Recipient #0
                Display name: "Andrei Ghita"
                Address type: "EX"
                Address: "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=24988e4426a04bceb9df2a1b61c23ff9-mailbox1"
                Tracking status: "0"
        Updating recipients...
        Saving changes...

Opening item with subject "test 2"
        Recipient #0
                Display name: "Andrei Ghita"
                Address type: "EX"
                Address: "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=24988e4426a04bceb9df2a1b61c23ff9-mailbox1"
                Tracking status: "3"
                ```
