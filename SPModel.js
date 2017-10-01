/*
    SPModel.js (Version 1.0)
    Contains class (including methods) to communicate with SharePoint using REST APIs or ESOM
    Author: Rohit Kumar Mighwal
    License: MIT
    Pre-requities: jQuery (min. ver 1.11)
    https://github.com/rkmighwal/SPModel.js
*/

// SharePoint Model class
var SPModel = function () {

    /*----- Private Members -----*/

    var _fileName = "SPModel.js",
        _scriptLoaded = false;

    var _sendAJAXCall = function (callURL, callback) {
        return jQuery.ajax({
            url: callURL,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" },
            success: callback,
            error: helper.logError
        });
    };

    // Method to convert file into array buffer
    var _getFileBuffer = function (file) {
        var deferred = jQuery.Deferred(),
            reader = new FileReader();

        reader.onload = function (e) {
            deferred.resolve(e.target.result);
        };

        reader.onerror = function (e) {
            deferred.reject(e.target.error);
        };

        reader.readAsArrayBuffer(file);

        return deferred.promise();
    };

    // Method to upload file as SharePoint list item attachment
    var _uploadFile = function (siteURL, listName, listItemId, fileName, file, attachments, callback) {
        _uploadFileSP(siteURL, listName, listItemId, fileName, file)
            .then(function (files) {
                attachments.splice(0, 1); // Remove first attachment from array
                _addAttachments(siteURL, listName, listItemId, attachments, callback); // Upload next attachment
            },
            function (sender, args) {
                helper.logError(args.get_message());
            });
    };

    // Method to upload file into SharePoint
    var _uploadFileSP = function (siteURL, listName, id, fileName, file) {
        var deferred = jQuery.Deferred();

        _getFileBuffer(file).then(
            function (buffer) {
                var bytes = new Uint8Array(buffer),
                    content = new SP.Base64EncodedByteArray(),
                    binary = '';

                for (var b = 0; b < bytes.length; b++)
                    binary += String.fromCharCode(bytes[b]);

                var scriptbase = siteURL + "/_layouts/15/",
                    _success = function () {
                        var createitem = new SP.RequestExecutor(siteURL),
                            restApiURL = "/_api/web/lists/GetByTitle('" + listName + "')/items(" + id + ")/AttachmentFiles/add(FileName='" + file.name + "')",
                            callURL = (siteURL === "/" ? "" : siteURL) + restApiURL; // Add Site URL with REST API URL

                        _scriptLoaded = true; // flag set to script is already loaded

                        createitem.executeAsync({
                            url: callURL,
                            method: "POST",
                            binaryStringRequestBody: true,
                            body: binary,
                            success: function () { deferred.resolve(); },
                            error: helper.logError,
                            state: "Update"
                        });
                    };

                if (!_scriptLoaded) {
                    jQuery.cachedScript(scriptbase + "SP.RequestExecutor.js").done(_success).fail(function (xhr, status, error) {
                        helper.logError(error);
                    });
                } else
                    _success();
            },
            function (error) {
                deferred.reject(error);
            }
        );

        return deferred.promise();
    };

    // Method to get list items from SharePoint List using REST API
    var _getListItems = function (siteURL, listName, itemId, queryString, callback) {
        if (helper.checkJQuery && listName && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            siteURL = siteURL || _spPageContextInfo.webServerRelativeUrl; // set default Site URL

            var restApiURL = "/_api/web/lists/GetByTitle('" + listName + "')/items" + (itemId ? "(" + itemId + ")" : ""),
                callURL = (siteURL === "/" ? "" : siteURL) + restApiURL, // Add Site URL with REST API URL
                listItemsCallbackData = [], // Used to store data between multiple AJAX calls
                itemsLimit = null, // Store items limit
                handleCallback = function (data) {
                    if (typeof data != "undefined" && data !== null) {
                        if (data.d.hasOwnProperty("results") && data.d.results !== null && data.d.results.length > 0) { // Process multiple items
                            jQuery.each(data.d.results, function (index, item) { // Traverse & store data
                                listItemsCallbackData.push(item);
                            });

                            if ((!itemsLimit || (!!itemsLimit && listItemsCallbackData.length < parseInt(itemsLimit))) && data.d.__next && data.d.__next !== "") {
                                _sendAJAXCall(data.d.__next, handleCallback);
                                return false;
                            } else {

                                if (!!itemsLimit) {
                                    itemsLimit = parseInt(itemsLimit); // Parse to Number

                                    if (listItemsCallbackData.length > itemsLimit) // Remove exceeded items
                                        listItemsCallbackData.splice(itemsLimit, (listItemsCallbackData.length - itemsLimit));
                                }

                                callback(listItemsCallbackData);
                            }
                        } else // Return single item
                            callback(data.d);
                    } else
                        callback(null); // Send null, if no data exists
                };

            if (queryString !== null && typeof queryString == "string") {
                callURL += (queryString.indexOf('?') === 0 ? "" : "?") + queryString; // add query string contains filter, select and other REST parameters
                itemsLimit = helper.getURLParameterValueByName("\\$top", callURL);
            }

            _sendAJAXCall(callURL, handleCallback); // Send AJAX call, and return promise
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to get list items count from SharePoint List using REST API
    var _getListItemCount = function (siteURL, listName, callback) {
        if (helper.checkJQuery && listName && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            siteURL = siteURL || _spPageContextInfo.webServerRelativeUrl; // set default Site URL

            var restApiURL = "/_api/web/lists/GetByTitle('" + listName + "')/itemcount",
                callURL = (siteURL === "/" ? "" : siteURL) + restApiURL, // Add Site URL with REST API URL
                handleCallback = function (data) {
                    if (typeof data != "undefined" && data !== null) {
                        callback(data.d);
                    } else
                        callback(null); // Send null, if no data exists
                };

            _sendAJAXCall(callURL, handleCallback); // Send AJAX call, and return promise
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to get current user groups from SharePoint using REST API
    var _getCurrentUserGroups = function (siteURL, callback) {
        if (helper.checkJQuery && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            siteURL = siteURL || _spPageContextInfo.webServerRelativeUrl; // set default Site URL

            var restApiURL = "/_api/web/currentUser?$select=Groups/Id&$expand=Groups/Id",
                callURL = (siteURL === "/" ? "" : siteURL) + restApiURL, // Add Site URL with REST API URL
                listItemsCallbackData = [], // Used to store data between multiple AJAX calls
                handleCallback = function (data) {
                    if (typeof data != "undefined" && data !== null) {
                        if (data.d.hasOwnProperty("Groups") && data.d.Groups !== null && data.d.Groups.results.length > 0) { // Process multiple items
                            jQuery.each(data.d.Groups.results, function (index, item) { // Traverse & store data
                                listItemsCallbackData.push(item.Id);
                            });
                        }

                        listItemsCallbackData.push(_spPageContextInfo.userId);

                        callback(listItemsCallbackData);
                    } else
                        callback(null); // Send null, if no data exists
                };

            _sendAJAXCall(callURL, handleCallback); // Send AJAX call, and return promise
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to get everyone user id from SharePoint using REST API
    var _getEveryoneUserId = function (siteURL, callback) {
        if (helper.checkJQuery && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            siteURL = siteURL || _spPageContextInfo.webServerRelativeUrl; // set default Site URL

            var restApiURL = "/_api/web/siteusers/getbyloginname(@v)?@v='c:0(.s|true'&$select=Id",
                callURL = (siteURL === "/" ? "" : siteURL) + restApiURL, // Add Site URL with REST API URL
                handleCallback = function (data) {
                    if (typeof data != "undefined" && data !== null && data.d.hasOwnProperty("Id"))
                        callback(data.d.Id);
                    else
                        callback(null); // Send null, if no data exists
                };

            _sendAJAXCall(callURL, handleCallback); // Send AJAX call, and return promise
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to save list item into SharePoint list using JSOM
    var _saveListItem = function (siteURL, listName, data, callback) {
        if (helper.checkJQuery && listName && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback)) && data) {
            var clientContext = siteURL ? (new SP.ClientContext(siteURL)) : (new SP.ClientContext.get_current()),
                list = clientContext.get_web().get_lists().getByTitle(listName),
                listItems = [];

            jQuery.each(data, function (index, item) { // Traverse & add items
                var itemCreateInfo = new SP.ListItemCreationInformation(),
                    listItem = list.addItem(itemCreateInfo);

                for (var property in item) {
                    if (item.hasOwnProperty(property) && !jQuery.isFunction(item[property]))
                        listItem.set_item(property, item[property]); // Set field values
                }

                listItem.update();
                clientContext.load(listItem);

                listItems.push(listItem);
            });

            clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
                callback(listItems);
            }), Function.createDelegate(this, function (sender, args) {
                helper.logError(args.get_message() + '\n' + args.get_stackTrace());
            }));
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to update list item into SharePoint list using JSOM
    var _updateListItem = function (siteURL, listName, data, callback) {
        if (helper.checkJQuery && listName && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback)) && (data && jQuery.isArray(data))) {
            var clientContext = siteURL ? (new SP.ClientContext(siteURL)) : (new SP.ClientContext.get_current()),
                list = clientContext.get_web().get_lists().getByTitle(listName),
                listItems = [];

            jQuery.each(data, function (index, item) { // Traverse & add items
                if (item.hasOwnProperty("ID") && item.hasOwnProperty("data") && (item.ID && (typeof item.ID == "number" || typeof item.ID == "string")) && item.data) {
                    var listItem = list.getItemById(parseInt(item.ID).toString()),
                        itemData = item.data;

                    for (var property in itemData) {
                        if (itemData.hasOwnProperty(property) && !jQuery.isFunction(itemData[property]))
                            listItem.set_item(property, itemData[property]); // Set field values
                    }

                    listItem.update();

                    clientContext.load(listItem);

                    listItems.push(listItem);
                }
            });

            clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
                callback(listItems);
            }), Function.createDelegate(this, function (sender, args) {
                helper.logError(args.get_message() + '\n' + args.get_stackTrace());
            }));
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to delete list item from SharePoint list using JSOM
    var _deleteListItem = function (siteURL, listName, data, callback) {
        if (helper.checkJQuery && listName && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback)) && (data && jQuery.isArray(data))) {
            var clientContext = siteURL ? (new SP.ClientContext(siteURL)) : (new SP.ClientContext.get_current()),
                list = clientContext.get_web().get_lists().getByTitle(listName);

            jQuery.each(data, function (index, item) { // Traverse & add items
                if (item.hasOwnProperty("ID") && (item.ID && (typeof item.ID == "number" || typeof item.ID == "string"))) {
                    var listItem = list.getItemById(parseInt(item.ID).toString());

                    listItem.deleteObject();
                }
            });

            clientContext.executeQueryAsync(callback, Function.createDelegate(this, function (sender, args) {
                helper.logError(args.get_message() + '\n' + args.get_stackTrace());
            }));
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to add multiple attachments
    var _addAttachments = function (siteURL, listName, itemId, attachments, callback) {
        if (helper.checkJQuery && listName && itemId && (attachments && jQuery.isArray(attachments)) && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            siteURL = siteURL || _spPageContextInfo.webServerRelativeUrl; // set default Site URL

            if (attachments.length > 0) {
                var attachment = attachments[0],
                    filename = attachment.name,
                    file = attachment.file;

                // Upload file as attachment
                _uploadFile(siteURL, listName, itemId, filename, file, attachments, callback);
            }
            else
                callback();
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to remove attachment from list item
    var _removeAttachment = function (siteURL, listName, listItemId, item) {
        var callUrl = siteURL + "/_api/lists/getByTitle('" + listName + "')/getItemById(" + listItemId + ")/AttachmentFiles/getByFileName(@F)?@F='" + item.name + "'",
            call = jQuery.ajax({
                url: callUrl,
                method: "DELETE",
                headers: { 'X-RequestDigest': jQuery('#__REQUESTDIGEST').val() }
            });

        return call;
    };

    // Method to ensure folder exists in SharePoint
    var _ensureFolder = function (siteURL, listName, listPath, folderName, callback) {
        if (helper.checkJQuery && listName && listPath && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            var clientContext = siteURL ? (new SP.ClientContext(siteURL)) : (new SP.ClientContext.get_current()),
                folderServerRelativePath = (listPath + "/" + folderName),
                folder = clientContext.get_web().getFolderByServerRelativeUrl(folderServerRelativePath);

            clientContext.load(folder);

            clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
                callback(true);
            }), Function.createDelegate(this, function (sender, args) {
                if (args.get_errorTypeName() === "System.IO.FileNotFoundException")
                    _createFolder(siteURL, listName, folderName, callback);
                else
                    helper.logError(args.get_message() + '\n' + args.get_stackTrace());
            }));
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to add folder under list/library
    var _createFolder = function (siteURL, listName, folderName, callback) {
        if (helper.checkJQuery && listName && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            var clientContext = siteURL ? (new SP.ClientContext(siteURL)) : (new SP.ClientContext.get_current()),
                list = clientContext.get_web().get_lists().getByTitle(listName);


            var folderCreationInfo = new SP.ListItemCreationInformation();
            folderCreationInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
            folderCreationInfo.set_leafName(folderName);

            var listItem = list.addItem(folderCreationInfo);
            listItem.update();

            clientContext.load(listItem);

            clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
                callback(true);
            }), Function.createDelegate(this, function (sender, args) {
                helper.logError(args.get_message() + '\n' + args.get_stackTrace());
            }));
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to add multiple documents in SharePoint List/Library
    var _addDocuments = function (siteURL, listName, listPath, folderName, documents, callback) {
        if (helper.checkJQuery && listName && listPath && folderName && (documents && jQuery.isArray(documents)) && (typeof callback != "undefined" && callback !== null && jQuery.isFunction(callback))) {
            siteURL = siteURL || _spPageContextInfo.webServerRelativeUrl; // set default Site URL

            if (documents.length > 0) {
                var document = documents[0],
                    filename = document.name,
                    file = document.file,
                    fileData = document.data || null;

                // Upload document to SharePoint
                _uploadDocument(siteURL, listName, listPath, folderName, filename, file, fileData, documents, callback);
            }
            else
                callback();
        } else
            throw new SyntaxError("Invalid Parameters.", _fileName);
    };

    // Method to upload file as SharePoint document
    var _uploadDocument = function (siteURL, listName, listPath, folderName, fileName, file, fileData, documents, callback) {
        _uploadDocumentSP(siteURL, listName, listPath, folderName, fileName, file, fileData)
            .then(function (files) {
                documents.splice(0, 1); // Remove first document from array
                _addDocuments(siteURL, listName, listPath, folderName, documents, callback); // Upload next document
            },
            function (sender, args) {
                helper.logError(args.get_message());
            });
    };

    // Method to upload file into SharePoint
    var _uploadDocumentSP = function (siteURL, listName, listPath, folderName, fileName, file, fileData) {
        var deferred = jQuery.Deferred();

        _getFileBuffer(file).then(
            function (buffer) {
                var bytes = new Uint8Array(buffer),
                    content = new SP.Base64EncodedByteArray(),
                    binary = '',
                    folderServerRelativePath = (listPath + "/" + folderName);

                for (var b = 0; b < bytes.length; b++)
                    binary += String.fromCharCode(bytes[b]);

                var scriptbase = siteURL + "/_layouts/15/",
                    _success = function () {
                        var createitem = new SP.RequestExecutor(siteURL),
                            restApiURL = "/_api/web/GetFolderByServerRelativeUrl('" + folderServerRelativePath + "')/files/add(overwrite=true, url='" + file.name + "')",
                            callURL = (siteURL === "/" ? "" : siteURL) + restApiURL; // Add Site URL with REST API URL

                        _scriptLoaded = true; // flag set to script is already loaded

                        createitem.executeAsync({
                            url: callURL,
                            method: "POST",
                            binaryStringRequestBody: true,
                            processData: false,
                            headers: {
                                "accept": "application/json;odata=verbose"
                            },
                            body: binary,
                            success: function (response) {
                                if (typeof fileData != "undefined" && fileData !== null && response && response.body) {
                                    var responseData = JSON.parse(response.body);

                                    _getFileListItem(responseData.d.ListItemAllFields.__deferred.uri).done(function (listItem) {
                                        _updateFileData(siteURL, listName, listItem.d, fileData, deferred);
                                    });
                                } else
                                    deferred.resolve();
                            },
                            error: helper.logError,
                            state: "Update"
                        });
                    };

                if (!_scriptLoaded) {
                    jQuery.cachedScript(scriptbase + "SP.RequestExecutor.js").done(_success).fail(function (xhr, status, error) {
                        helper.logError(error);
                    });
                } else
                    _success();
            },
            function (error) {
                deferred.reject(error);
            }
        );

        return deferred.promise();
    };

    // Method to get list item related to file URI
    var _getFileListItem = function (fileUri) {
        return jQuery.ajax({
            url: fileUri,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" }
        });
    };

    // Method to update file data, after file created
    var _updateFileData = function (siteURL, listName, listItemData, fileData, deferred) {
        var itemData = {
            ID: listItemData.ID,
            data: fileData
        };

        _updateListItem(siteURL, listName, [itemData], function () {
            deferred.resolve();
        });
    };

    /*----- Public Members -----*/

    // Method to get list items from SharePoint List using JSOM
    this.getListItems = _getListItems;
    this.saveListItem = _saveListItem;
    this.updateListItem = _updateListItem;
    this.getListItemCount = _getListItemCount;
    this.deleteListItem = _deleteListItem;
    this.getCurrentUserGroups = _getCurrentUserGroups;
    this.getEveryoneUserId = _getEveryoneUserId;
    this.addAttachments = _addAttachments;
    this.removeAttachment = _removeAttachment;
    this.ensureFolder = _ensureFolder;
    this.addDocuments = _addDocuments;
};