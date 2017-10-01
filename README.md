# SPModel.js
JavaScript Library contains different methods for communicate with SharePoint lists/libraries.

Having following methods:

getListItems - Get List Items from SharePoint, using REST API, can pass REST Filters and Number of items required (optional)

Syntax
var _spModel = new SPModel();
_spModel.getListItems(Site URL, List Name, List Item ID, REST API Parameters, Callback Method);

Parameters
Site URL (required) - SharePoint Site URL (Must have appropriate permissions), default value - Current SharePoint Site
List Name (required) - SharePoint List/Library Name
List Item ID (optional, pass null if not required) - Pass Item ID to get details for particular list item
REST API Parameters (optional, pass null if not required) - OData parameters to filter output (using $filter), Limit items count (using $top), or Select specific fields (using $select), Please refer MSDN for more details (https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/use-odata-query-operations-in-sharepoint-rest-requests)
Callback Method (required) - Must be a valid JavaScript function, List item data will be passed as an array of objects to callback method.

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
