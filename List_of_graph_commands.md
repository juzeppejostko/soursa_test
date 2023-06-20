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

### List messages in folder, but show only specified parameters
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
https://graph.microsoft.com/v1.0/me/messages?($filter=ReceivedDateTime ge 2023-06-18 and ReceivedDateTime le 2023-06-19)&($count=true)

### Get messages with the specific subject
https://graph.microsoft.com/v1.0/me/messages?$filter=subject eq 'Fwd: 1039996 // Grotex_company details for BAFA'

### Use multiple queries in one request
Use brackets to combine multiple queries in one request
https://graph.microsoft.com/v1.0/me/messages?($select=parentFolderId,id,createdDateTime,sentDateTime,sender,from,toRecipients,ccRecipients,subject)($filter=subject eq 'Fwd: 1039996 // Grotex_company details for BAFA')


### Get messages from specific period
https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$orderby=receivedDateTime desc&$filter=receivedDateTime ge 2023-06-10T00:00Z

Get messages from April, 2023
https://graph.microsoft.com/v1.0/me/messages?$filter=ReceivedDateTime ge 2023-04-01 and receivedDateTime lt 2023-05-01

ge - means greater or equal, full list is avasilible here: https://learn.microsoft.com/en-us/graph/filter-query-parameter?tabs=http

### Quantity of sent messages
https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages?$count=true


## Examples
https://github.com/microsoftgraph/msgraph-sdk-javascript/tree/dev/samples/javascript
https://github.com/microsoftgraph/msgraph-sample-javascriptspa
https://github.com/microsoftgraph/msgraph-sample-office-addin
https://github.com/microsoftgraph/msgraph-sample-reactspa


https://learn.microsoft.com/ru-ru/graph/tutorials/javascript?tabs=aad