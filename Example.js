/*
    SPModel.js Sample
*/

var _spModel = new SPModel(); // object of Model class
    
var apiParameters = "$select=Id,Title&$filter=Title eq 'Test'";
_spModel.getListItems(_spPageContextInfo.webServerRelativeUrl, 'Test', null, apiParameters, getListItemsCallback); // Call method to get data from SharePoint
    
var getListItemsCallback = function (data) {
    if (typeof data !== "undefined" && ((Array.isArray(data) && data.length > 0) || data !== null)) {
        jQuery.each(data, function (index, item) { // Process all result data
            if (item !== null && item.Title) {
                console.log('ID: ' + item.Id + ', Title: ' + item.Title);
            }
        });
    }
};
    
var getEveryOneuserId = function () {
    _spModel.getEveryoneUserId(_spPageContextInfo.webServerRelativeUrl, processEveryoneUserId); // Call method to get data from SharePoint
};
    
var processEveryoneUserId = function (data) {
    if (typeof data !== "undefined" && data !== null) {
        console.log(data); // Get value from SP
    }
};
    
var getCurrentUserGroupIds = function () {
    _spModel.getCurrentUserGroups(null, processCurrentUserGroupIds); // Call method to get data from SharePoint
};
    
var processCurrentUserGroupIds = function (data) {
    if (typeof data !== "undefined" && ((Array.isArray(data) && data.length > 0) || data !== null)) {
        jQuery.each(data, function (index, item) { // Process all result data
            if (item !== null && item.Title) {
                console.log('ID: ' + item.Id + ', Title: ' + item.Title);
            }
        });
    }
};
    
var getListItemsCount = function () {
    _spModel.getListItemCount(null, 'Test', processReportsCount); // Call method to get data from SharePoint
};

var processReportsCount = function (data) {
    if (typeof data !== "undefined" && ((Array.isArray(data) && data.length > 0) || data !== null))
        console.log(parseInt(data.ItemCount));
};

var _saveItem = function () {
        var formData = {
            Title: 'Test'
        };

        if (formData !== null) {
            if (_view.itemID) {
                var _temp = formData;

                formData = {
                    ID: 1, // Item ID to update existing item
                    data: _temp
                };

                _spModel.updateListItem(null, 'Test', [formData], function (listItems) {
                    if (listItems && Array.isArray(listItems) && listItems.length > 0) {
                        listItems = listItems[0];

                        console.log(listItems.get_id());
                    }
                }); // Call method to save data to SharePoint
            } else {
                _spModel.saveListItem(null, 'Test', [formData], function (listItems) {
                    if (listItems && Array.isArray(listItems) && listItems.length > 0) {
                        listItems = listItems[0];

                        console.log(listItems.get_id());
                    }
                }); // Call method to save data to SharePoint
            }
        }
};
    
// Method to remove existing attachments, if any
// And add new attachments, if any

var attachment = function () { // Attachment class
    this.name = "";
    this.file = null;
};

var _updateAttachments = function () {
    var attachments = {
        removed: [],
        added: []
    }

    // Add new attachment to array for adding to list item, need file content and file name both for adding
    var _attachment = new attachment();

    _attachment.name = "Test.txt";
    _attachment.file = jQuery('input[type="file"]').get(0).files[0];

    attachments.added.push(_attachment);

    // Add new attachment to array for removing from list item, just need file name to be removed
    var _attachment = new attachment();
    
    _attachment.name = "Test.txt";

    attachments.removed.push(_attachment);

    if (attachments.removed.length > 0) {
        var calls = new Array();

        jQuery.each(attachments.removed, function (index, item) {
            calls.push(_spModel.removeAttachment(null, 'Test', 1, item));
        });

        jQuery.when.apply(null, calls).done(function () {
            _spModel.addAttachments(null, 'Test', 1, attachments.added, function () {
                console.log('Attachments Added');
            });
        });

    } else
        _spModel.addAttachments(null, 'Test', 1, attachments.added, function () {
            console.log('Attachments Added');
        });
};