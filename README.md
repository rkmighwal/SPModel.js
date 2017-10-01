# SPModel.js
JavaScript Library contains different methods for communicate with SharePoint lists/libraries.

Please refer Example.js for sample method calls. Log a issue for any query :)

Having following methods:

getListItems - Get List Items from SharePoint, using REST API, can pass REST Filters and Number of items required (optional)

saveListItem - Save List Item to SharePoint list, using SharePoint ESOM 
    
updateListItem - Update existing List Item in SharePoint list/library, using SharePoint ESOM

getListItemCount - Get Count of List Items in a SharePoint list, using REST API

deleteListItem - Delete a List Item (with List Item ID) from SharePoint List/Library, using SharePoint ESOM

getCurrentUserGroups - Get All SharePoint Groups of Current User, using SharePoint REST API

getEveryoneUserId - Get User ID for Everyone User in SharePoint, using SharePoint REST API

addAttachments - Add Attachments to a SharePoint List Item, using SharePoint REST API

removeAttachment - Remove Attachment (using File Name) from SharePoint List Item, using SharePoint REST API

ensureFolder - Ensure a folder exists in SharePoint Library, using SharePoint ESOM

addDocuments - Add Documents to SharePoint Library, using SharePoint REST API
