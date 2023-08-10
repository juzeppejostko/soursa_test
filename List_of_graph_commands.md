## Common commands for graph

Link on Microsoft Graph exlorer: https://developer.microsoft.com/en-us/graph/graph-explorer

### List of folders:
https://graph.microsoft.com/v1.0/me/mailFolders/

### List messages in folder:
GET /me/mailFolders/{id}/messages

or

https://graph.microsoft.com/v1.0/me/mailFolders/"AAMkADQ0NDQ3MjIxLTgzYjAtNDQ2Yy04ODJlLTEyODVhMjdlMDkxNQAuAAAAAAAeB4YSD81iRYQYxh-eLO7mAQBCSp8yCjLkTKJp2P4zTumQAAAAAAEJAAA="/messages

or same but skipping first 10 items

https://graph.microsoft.com/v1.0/me/mailfolders/"AAMkADQ0NDQ3MjIxLTgzYjAtNDQ2Yy04ODJlLTEyODVhMjdlMDkxNQAuAAAAAAAeB4YSD81iRYQYxh-eLO7mAQBCSp8yCjLkTKJp2P4zTumQAAAB374FAAA="/messages?$skip=10

### List messages in folder, but show only specified parameters ('select' parameter)
https://graph.microsoft.com/v1.0/me/messages?$select=parentFolderId,id,createdDateTime,sentDateTime,sender,from,toRecipients,ccRecipients,subject,body

### List childFolders of folder:
GET /me/mailFolders/{id}/childFolders

or

https://graph.microsoft.com/v1.0/me/mailfolders/"AAMkADQ0NDQ3MjIxLTgzYjAtNDQ2Yy04ODJlLTEyODVhMjdlMDkxNQAuAAAAAAAeB4YSD81iRYQYxh-eLO7mAQBCSp8yCjLkTKJp2P4zTumQAAAAAGAoAAA="/childFolders

### Get a specific message
GET https://graph.microsoft.com/v1.0/me/messages/AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OABGAAAAAAAiQ8W967B7TKBjgx9rVEURBwAiIsqMbYjsT5e-T7KzowPTAAAAAAEMAAAiIsqMbYjsT5e-T7KzowPTAASoXUT3AAA=

https://learn.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http

### Get a number of objects (messages)
add '?$count=true'at the end of query

https://graph.microsoft.com/v1.0/me/mailfolders/"AAMkADQ0NDQ3MjIxLTgzYjAtNDQ2Yy04ODJlLTEyODVhMjdlMDkxNQAuAAAAAAAeB4YSD81iRYQYxh-eLO7mAQBCSp8yCjLkTKJp2P4zTumQAAAB374FAAA="/messages?$count=true

!!!IMPORTANT
'$count=true' doesn't work in queries i.e. next request would not count qty of elements:
<p><a href=""><del>https://graph.microsoft.com/v1.0/me/messages?($filter=ReceivedDateTime ge 2023-06-18 and ReceivedDateTime le 2023-06-19)&($count=true)</del></a></p>

### Get messages with the specific subject
<p><a href="#">https://graph.microsoft.com/v1.0/me/messages?$filter=subject eq 'Fwd: 1039202 // Index_company details'</a></p>

### Use multiple queries in one request
Use brackets to combine multiple queries in one request<br>
<p><a href="#">https://graph.microsoft.com/v1.0/me/messages?($select=parentFolderId,id,createdDateTime,sentDateTime,sender,from,toRecipients,ccRecipients,subject)($filter=subject eq 'Fwd: 1039202 // Index_company details')</a></p>


### Get messages from specific period
<p><a href="#">https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$orderby=receivedDateTime desc&$filter=receivedDateTime ge 2023-06-10T00:00Z</a></p>

Get messages before 7th of May 2023
<p><a href="#">graph.microsoft.com/me/messages?$filter=receivedDateTime lt 2023-05-07T16:55:39Z&$orderby=receivedDateTime desc</a></p>

<b>Get all messages with the recivedeDateTime May 2023</b><br>
<p><a href="#">https://graph.microsoft.com/v1.0/me/messages/?$orderby=receivedDateTime desc &$filter=receivedDateTime ge 2019-01-01T00:00:00Z and receivedDateTime lt 2023-08-01T00:00:00Z &$select=receivedDateTime,webLink,sender,from,toRecipients,ccRecipients&$count=true</a></p>

ge - means greater or equal, full list is avasilible here: https://learn.microsoft.com/en-us/graph/filter-query-parameter?tabs=http

<p>'$count=true'here does not work</p>

### Quantity of sent messages
https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages?$count=true


## Examples
https://github.com/microsoftgraph/msgraph-sdk-javascript/tree/dev/samples/javascript
https://github.com/microsoftgraph/msgraph-sample-javascriptspa
https://github.com/microsoftgraph/msgraph-sample-office-addin
https://github.com/microsoftgraph/msgraph-sample-reactspa


https://learn.microsoft.com/ru-ru/graph/tutorials/javascript?tabs=aad