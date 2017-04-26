/**
 * Author: Juan M Alcantara Acosta jmalcantara1@gmail.com
 * with modifications from google scripts public functions
 * v2.22 compatible with DHIS2 2.22-2.24
 *
 * Retrieves all the rows in the active spreadsheet that contain data and creates an array, indexed by its normalized column name.
 * Mandatory columns:
 * parentcode or parentid, name, shortname, code
 * Optional columns:
 * id, openingdate, comment, latitude, longitude, contactperson, address, email, phonenumber, description
 *
 * The function requires the headers in the first row, starting with the first column (A)
 * headers = A1, B1, C1...
 * data = A2, B2, C2...
 *
 *
 * To reference the different objects use the normalized column names.
 *  Example:
 *  Parent_Code -> parentcode
 *  Parent Code -> parentCode
 *
 */

//
// Create one JSON file with all user accounts
//


function createSingleJSON() {
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numColumns = range.getNumColumns();
    var spreadsheetDataRange = sheet.getRange(2, range.getColumn(), numRows - 1, numColumns);
    var spreadsheetDataRows = spreadsheetDataRange.getNumRows();

    // Create objects from rows
    var spreadsheetObjects = getRowsData(sheet, spreadsheetDataRange);

    // A variable can be used to assigne the values of a specific object and remove the array member number
    // var oneSpreadsheetObject = spreadsheetObjects[0]
    // Logger.log("Object: " + oneSpreadsheetObject.name);
    //
    //
    var users = 0;
    var skipped = 0;
    var jsondoc = '{'+'\"users\"'+':[';
    for (var i = 0; i <= (spreadsheetObjects.length - 1); i++) {
        if (spreadsheetObjects[i].emailaddress != undefined && spreadsheetObjects[i].json != undefined && spreadsheetObjects[i].username != undefined && spreadsheetObjects[i].configuredaccounts != 'ok') {
            if (spreadsheetObjects[i].emailaddress != 'no email') {
                // Assamble JSON here
                //
                // jsondoc = '{ "email": "' + spreadsheetObjects[i].emailaddress + '"'
                // jsondoc += ','
                // , "organisationUnits": [{"id":"V4YMV9ds7Gm"},{"id":"L8MKd63FAn7"},{"id":"f3AmtnHWgRe"},{"id":"KTI1PO5DGM6"},{"id":"hhWN4I5z7mY"}]
                // ,"dataViewOrganisationUnits": [{"id":"V4YMV9ds7Gm"},{"id":"L8MKd63FAn7"},{"id":"f3AmtnHWgRe"},{"id":"KTI1PO5DGM6"},{"id":"hhWN4I5z7mY"}]
                // , "groups": [{ "id": "pFx6SYn4lot"},{ "id": "guH0d8sZ7SW"}]
                // jsondoc += ', "firstName":"' + spreadsheetObjects[i].firstname + '","surname":"' + spreadsheetObjects[i].surname + '","employer":"' + spreadsheetObjects[i].employer + '"'
                // ,"userCredentials":{"username":"mgizachew","userAuthorityGroups":[{"id":"Cacew7RnCqe"},{"id":"QvMVe4QPrzY"},{"id":"o25FaENe7Hp"},{"id":"kAXajGyejhd"},{"id":"tKtr4n0xZTE"}]} }
                //
                // End assambling here

                if (users > 0)
                {
                    jsondoc += ',\n';
                }
                jsondoc += spreadsheetObjects[i].json;
                users++;
                Logger.log("Record: " + users);

            } else
            {
                skipped++;
            }

        } else
        {
            skipped++;
        }
    }
    jsondoc += '\n]}'
    Logger.log(users);
    Logger.log(jsondoc)

    var strDocName = sheet.getName().replace(/\s/g, '') + '_' + users + '_users_' + Utilities.formatDate(new Date(), 'GMT', "yyyy-MM-dd'T'HH-mm");
    DriveApp.createFile(strDocName + '.json', jsondoc, 'application/txt');
    SpreadsheetApp.getActiveSpreadsheet().toast("Document created: " + users + " users, Records skipped: " + skipped);
    Logger.log('Skipped records ' + skipped);

};

// ------------------------------------------------------------------------------------------------
//  Post and file functions - User Groups and User Settings
// ------------------------------------------------------------------------------------------------


function XMLuserGroups()
{
    createUserGroups();
};

function postUserGroups()
{

    var task = 'userGroups';
    var post = true;
    showPasswordDialog( task, post );

};

function postUserSettings()
{

    var task = "userSettings";
    showPasswordDialog( task );

};


// ------------------------------------------------------------------------------------------------
//  Config function - Get config information
// ------------------------------------------------------------------------------------------------

function getConfig()
{
// Process Configuration

    var sheet = spreadsheet.getSheetByName('Config');
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numColumns = range.getNumColumns();
    var ObjectsDataRange = sheet.getRange(3, range.getColumn(), numRows - 1, numColumns);
    var numObjectsDataRows = ObjectsDataRange.getNumRows() - 1;


    // Create objects from rows
    var ConfigObjects = getRowsData(sheet, ObjectsDataRange, 2);
    return ConfigObjects;

}


// ------------------------------------------------------------------------------------------------
// Add existing user accounts to user groups. The function doesn't verify if the user exists, the system
// will return an error when the account doesn't exist.
// ------------------------------------------------------------------------------------------------

function createUserGroups(username,  userpassword, post) {
    var spreadsheet = SpreadsheetApp.getActive();

    // Process Configuration

    var sheet = spreadsheet.getSheetByName('Config');
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numColumns = range.getNumColumns();
    var ObjectsDataRange = sheet.getRange(3, range.getColumn(), numRows - 1, numColumns);
    var numObjectsDataRows = ObjectsDataRange.getNumRows() - 1;


    // Create objects from rows
    var CodeObjects = getRowsData(sheet, ObjectsDataRange, 2);

    var users = 0;
    var groups = 0;
    var currentUser = null;
    if (CodeObjects[0].server != undefined) {
        // Check if the server URL is valid!
        //
        var url = CodeObjects[0].server;
        if (url.substring(url.length - 1,url.length) == "/") {
            url = url.substring(0,url.length - 1)
        }
    }
    else {
        var url = 'https://play.dhis2.org/demo';
    }

    Logger.log(groups);
    Logger.log('Post: ' + post);
    Logger.log('URL: ' + url);

    // Process main spreadsheet

    var sheet = SpreadsheetApp.getActiveSheet();
    //var sheet = spreadsheet.getSheetByName('New-Users');
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numColumns = range.getNumColumns();
    var spreadsheetDataRange = sheet.getRange(2, range.getColumn(), numRows - 1, numColumns);
    var spreadsheetDataRows = spreadsheetDataRange.getNumRows() - 1;

    // Create objects from rows
    var spreadsheetObjects = getRowsData(sheet, spreadsheetDataRange);

    var xmlUserGroups = XmlService.createElement('userGroups');

    SpreadsheetApp.getActiveSpreadsheet().toast( "Processing userGroups...");

    for (var i = 0; i < numObjectsDataRows; i++) {
        if (CodeObjects[i].groupname != undefined && CodeObjects[i].groupuid != undefined) {
            Logger.log(CodeObjects[i].groupname);
            Logger.log(CodeObjects[i].groupuid);
            currentGroupName = CodeObjects[i].groupname;
            currentGroupUID = CodeObjects[i].groupuid;

            var xmlUsersList = XmlService.createElement('users');

            for (var x = 0; x < spreadsheetDataRows; x++) {
                currentUID = (spreadsheetObjects[x].uid);
                currentRecord = spreadsheetObjects[x].configuredgroups;
                currentX = x;
                if ( spreadsheetObjects[x].uid != undefined && spreadsheetObjects[x].configuredgroups != "ok") {

                    if ( spreadsheetObjects[x].usergroup01 != undefined && currentUser == null ) {
                        if ( spreadsheetObjects[x].usergroup01 == CodeObjects[i].groupname ) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup02 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup02 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup03 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup03 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup04 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup04 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup05 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup05 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup06 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup06 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup07 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup07 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup08 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup08 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup09 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup09 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup10 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup10 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup11 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup11 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup12 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup12 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup13 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup13 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup14 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup14 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup15 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup15 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup16 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup16 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup17 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup17 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup18 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup18 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup19 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup19 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }
                    if (spreadsheetObjects[x].usergroup20 != undefined && currentUser == null ) {
                        if (spreadsheetObjects[x].usergroup20 == CodeObjects[i].groupname) {
                            users++;

                            var xmlUser = XmlService.createElement('user')
                                .setAttribute('id', spreadsheetObjects[x].uid);
                            xmlUsersList.addContent(xmlUser);
                            currentUser = spreadsheetObjects[x].uid;
                        }
                    }

                    if ( currentUser != null && post == 'true'  ) {
                        // Update the collection adding the user account
                        // Uses URL from Config spreadsheet to allow the user to change which server will be used
                        // /api/{collection-object}/{collection-object-id}/{collection-name}/{object-id}
                        var apiCall = url + "/api/userGroups/" + CodeObjects[i].groupuid  + "/users/" + currentUser;
                        var response = dhisAPIcall(apiCall,"post",username,userpassword);
                        // Log responses
                        Logger.log(response);
                    }
                    currentUser = null;
                }
            }

            if ( users > 0 ) {
                var xmlUsergroup = XmlService.createElement('userGroup')
                    .setAttribute('id', CodeObjects[i].groupuid)
                    .setAttribute('name', CodeObjects[i].groupname);
                xmlUsergroup.addContent(xmlUsersList);
                xmlUserGroups.addContent(xmlUsergroup);
            }
            Logger.log(users);
            users = 0;

        }
    }

    if ( post != 'true' ) {

        // Create XML document
        var url = XmlService.getNamespace('http://dhis2.org/schema/dxf/2.0');
        var xmlroot = XmlService.createElement('metaData',url);
        xmlroot.addContent(xmlUserGroups);
        var document = XmlService.createDocument(xmlroot);
        var xmldoc = XmlService.getPrettyFormat().format(document);
        Logger.log(xmldoc);

        //  Save the XML output to a document using the name of the active spreadsheet
        var strDocName = sheet.getName().replace(/\s/g, '') + '_userGroups_' + Utilities.formatDate(new Date(), 'GMT', "yyyy-MM-dd");
        Logger.log(sheet.getName());
        var folder = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
        Logger.log("folder:");
        Logger.log(folder.getName());
        DriveApp.createFile(strDocName + '.xml', xmldoc, 'application/xml');
        SpreadsheetApp.getActiveSpreadsheet().toast("Document created: " + strDocName + '.xml');
    }
    else {
        Logger.log('End');
        SpreadsheetApp.getActiveSpreadsheet().toast('Post to ' + url + ' completed');
    }

};

// ------------------------------------------------------------------------------------------------
// Process user settings for existing accounts. The function doesn't verify if the user exists, the system
// will return an error when the account doesn't exist.
// ------------------------------------------------------------------------------------------------

function userSettings(username,userpassword) {
    var spreadsheet = SpreadsheetApp.getActive();

    // Process Configuration

    var sheet = spreadsheet.getSheetByName('Config');
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numColumns = range.getNumColumns();
    var ObjectsDataRange = sheet.getRange(3, range.getColumn(), numRows - 1, numColumns);
    var numObjectsDataRows = ObjectsDataRange.getNumRows() - 1;

    // Create objects from rows
    var CodeObjects = getRowsData(sheet, ObjectsDataRange, 2);

    var users = 0;
    var groups = 0;
    var currentUser = null;

    Logger.log(groups);
    // Process main

    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numColumns = range.getNumColumns();
    var spreadsheetDataRange = sheet.getRange(2, range.getColumn(), numRows - 1, numColumns);
    var spreadsheetDataRows = spreadsheetDataRange.getNumRows();

    // Create objects from rows
    var spreadsheetObjects = getRowsData(sheet, spreadsheetDataRange);

    if (CodeObjects[0].server != undefined) {
        var url = CodeObjects[0].server;
        if (url.substring(url.length - 1,url.length) == "/") {
            url = url.substring(0,url.length - 1)
        }
        // Verify if url includes https:// or http://

    }
    else {
        var url = 'https://play.dhis2.org/demo';
    }

    var call = "";
    var response = "";

    SpreadsheetApp.getActiveSpreadsheet().toast( "Processing userSettings");

    for (var i = 0; i < spreadsheetDataRows; i++) {
        if ( spreadsheetObjects[i].username != undefined && spreadsheetObjects[i].uid != undefined && spreadsheetObjects[i].configuredsettings != "ok" ) {

            if ( spreadsheetObjects[i].uilanguage != undefined) {
                // /api/userSettings/my-key?user=username&value=my-val
                call = url + "/api/userSettings/keyUiLocale?user=" + spreadsheetObjects[i].username + "&value=" + spreadsheetObjects[i].uilanguage;
                Logger.log(call);
                response = dhisAPIcall(call,"POST",username,userpassword);
            }

            if ( spreadsheetObjects[i].dblanguage != undefined) {
                // /api/userSettings/my-key?user=username&value=my-val
                call = url + "/api/userSettings/keyDbLocale?user=" + spreadsheetObjects[i].username + "&value=" + spreadsheetObjects[i].dblanguage;
                Logger.log(call);
                response = dhisAPIcall(call,"POST",username,userpassword);
            }

        }
    }

};

//
//

function dhisAPIcall(call,method,user,password) {


    var credentials = user + ":" + password;
    var options =
        {
            "contentType" : "application/xml",
            "method" : method,
            "headers": {
                "Authorization": "Basic " + Utilities.base64Encode(credentials)
            }
        };
    try {

        var response = UrlFetchApp.fetch(call, options);
    } catch (e) {
        Logger.log("Error: " + e);
        SpreadsheetApp.getActiveSpreadsheet().toast("Error: "+ e);
        return e;
    }
    Logger.log(response);
    return response;
};




// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
    columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1 ; //the range should not include the header
    var numColumns = range.getLastColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
    var headers = headersRange.getValues()[0];
    Logger.log('Headers length: ' + headers.length);
    Logger.log('Headers: ' + headers);
    return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty(cellData)) {
                continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
    var keys = [];
    for (var i = 0; i < headers.length; ++i) {
        var key = normalizeHeader(headers[i]);
        if (key.length > 0) {
            keys.push(key);
            Logger.log(key);
        }
    }
    return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
//
// Normalization returns an all lowercase name if instead of spaces a symbol, hiphen or underdash is used.
//
// Examples:
//   "First_Name" -> "fistname"
//
function normalizeHeader(header) {
    var key = "";
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
        var letter = header[i];
        if (letter == " " && key.length > 0) {
            upperCase = true;
            continue;
        }
        if (!isAlnum(letter)) {
            continue;
        }
        if (key.length == 0 && isDigit(letter)) {
            continue; // first character must be a letter
        }
        if (upperCase) {
            upperCase = false;
            key += letter.toUpperCase();
        } else {
            key += letter.toLowerCase();
        }
    }
    return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
    return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
    return char >= 'A' && char <= 'Z' ||
        char >= 'a' && char <= 'z' ||
        isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
    return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
    if (data.length == 0 || data[0].length == 0) {
        return null;
    }

    var ret = [];
    for (var i = 0; i < data[0].length; ++i) {
        ret.push([]);
    }

    for (var i = 0; i < data.length; ++i) {
        for (var j = 0; j < data[i].length; ++j) {
            ret[j][i] = data[i][j];
        }
    }

    return ret;
}

// -----------------------------------------------------------------------------------------
// Password dialog box
// -----------------------------------------------------------------------------------------

function showPasswordDialog( task, post )
{
    var app = createPasswordDialog( task, post );

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    spreadsheet.show(app);
    return app;
}

// Create "Set Password" dialog
function createPasswordDialog( task, post )
{
    var app = UiApp.createApplication();
    app.setWidth(400);
    app.setHeight(100);

    var labelUsername = app.createLabel('Username');
    app.add(labelUsername);

    var usernameField = app.createTextBox().setId('usernameField').setName("username");
    app.add(usernameField);

    var labelPassword = app.createLabel('Password');
    app.add(labelPassword);

    var passwordField = app.createPasswordTextBox().setId('passwordField').setName("password");
    app.add(passwordField);

    var taskField = app.createTextBox().setId('taskField').setName("task");
    taskField.setVisible(false);
    taskField.setValue( task );
    app.add(taskField);

    var postField = app.createTextBox().setId('postField').setName("post");
    postField.setVisible(false);
    postField.setValue( post );
    app.add(postField);

    var handlerOkBtn = app.createServerHandler("setPassword")
        .addCallbackElement(passwordField)
        .addCallbackElement(usernameField)
        .addCallbackElement(taskField)
        .addCallbackElement(postField);
    var okBtn = app.createButton("OK", handlerOkBtn);
    okBtn.setWidth(60);
    app.add(okBtn);

    var handlerCancelBtn = app.createServerHandler("closePasswordDialog");
    var cancelBtn = app.createButton("Cancel", handlerCancelBtn);
    cancelBtn.setWidth(60);
    app.add( cancelBtn );

    var messageLbl = app.createLabel().setId("warningMessageLbl");
    messageLbl.setStyleAttribute( "color", "red" );
    app.add(messageLbl);

    return app;
}

// Set password
function setPassword( e )
{

    var app = UiApp.getActiveApplication();
    var username = e.parameter.username;
    var password = e.parameter.password;
    var task = e.parameter.task;
    var post = e.parameter.post;

    if( username == "" || password == "" )
    {
        app.getElementById("warningMessageLbl").setText("Please enter username/password ");
        return app;
    }
    else
    {
        app.getElementById("warningMessageLbl").setText("");
        if ( task == "userGroups" ) {
            app.getElementById("warningMessageLbl").setText("Processing userGroups...");
            createUserGroups( username, password, post );
        }
        else if ( task == "userSettings" ) {
            app.getElementById("warningMessageLbl").setText("Processing userSettings...");
            userSettings(username,password);
        }
        app.close();
        return app;


    }

}

// Close dialog
function closePasswordDialog() {
    var app = UiApp.getActiveApplication();
    app.close();
    return app;
}





/**
 * Adds a custom menu to the active spreadsheet.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [
        {
            name : 'Create Users (JSON)',
            functionName : 'createSingleJSON'
        },
        {
            name : 'Create userGroups (XML)',
            functionName : 'XMLuserGroups'
            //functionName : 'createUserGroups'
        },
        {
            name : 'Update userGroups (post)',
            functionName : 'postUserGroups'
            //functionName : 'createUserGroups'
        },
        {
            name : 'Update userSettings (post)',
            functionName : 'postUserSettings'
            //functionName : 'userSettings'
        }
    ];
    spreadsheet.addMenu('Script Menu', entries);
};
