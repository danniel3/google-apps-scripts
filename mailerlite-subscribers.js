function subscribers_pull() {

  // Define Sheet properties
  var ss = SpreadsheetApp.openById("ADD-SPREADSHEET-ID-HERE");
  var sheet = ss.setActiveSheet(ss.getSheetByName("ADD-SHEET-NAME-HERE"));

  var numRows = sheet.getLastRow();
  var startRow = 2;
  var startCol = 1;
  var numCols = 6;

  // Get total active subscribers

  var subscriberRange = sheet.getRange(startRow,startCol+1,numRows,1).getValues().filter(function(item){if (item[0] === 'active') {return true;} else {return false;}}).length-20;

  // Call Mailerlite API
  var url = "https://api.mailerlite.com/api/v2/subscribers?limit=1000&type=active&offset=" + subscriberRange;
  var headers = {
             "contentType": "application/json",
             "headers":{"X-MailerLite-ApiDocs": "true",
                        "X-MailerLite-ApiKey": "ADD-API-KEY-HERE",
                        "User-Agent": "ReadMe-API-Explorer"}
         };
  var response = UrlFetchApp.fetch(url, headers);
  var regex = response.toString().replace(/:\s*(-?\d+),/g, ': "$1",');
  var regarray = JSON.parse(regex);
  var newarray = Object.values(regarray);


  for (var obj in newarray){

    var field = newarray[obj];
    var nested = field.fields;
    var id = field.id;
    var type = field.type;
    var name = nested.find(element => element.key === "name").value;
    var email = nested.find(element => element.key === "email").value;
    var firstname = nested.find(element => element.key === "firstname").value;
    var lastname = nested.find(element => element.key === "last_name").value;

    // Set up first and last names
    if (firstname.length > 0){
      firstname = firstname;
    }
    else if (name.length < 3){
      firstname = name;
    }
    else{
      names = name.split(" ");
      firstname = names[0].replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
    }

    if (lastname.length > 0){
      lastname = lastname.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
    }
    else if (name.length < 3){
      lastname = "";
    }
    else{
      names = name.split(" ");
      lastname = name.replace(names[0],"").replace(/^ +/g,"").replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
    }


    // Update values if existing
    if (sheet.getRange(startRow,startCol,numRows).createTextFinder(id).findNext() !== null){
      sheet.getRange(sheet.getRange(startRow,startCol,numRows,1).createTextFinder(id).findNext().getRowIndex(),startCol,1,numCols).setValues([[id,type,name,email,firstname,lastname]]);
      Logger.log("Updated: " + id + " | Name: " + name + " | at " + email + ".");
    }

    // Add new row if non-existing
    else{
      sheet.getRange(numRows+1,startCol,1,numCols).setValues([[id,type,name,email,firstname,lastname]]);
      sheet.appendRow([""]);
      numRows++;
      Logger.log("Created: " + id + " | Name: " + name + " | at " + email + ".");
    }
  }
}

function subscribers_push() {

  // Define Sheet properties
  var ss = SpreadsheetApp.openById("ADD-SPREADSHEET-ID-HERE");
  var sheet = ss.setActiveSheet(ss.getSheetByName("ADD-SHEET-NAME-HERE"));
  var numRows = sheet.getLastRow();
  var startRow = numRows - 100;
  var startCol = 1;
  var numCols = 6;

  // Fetch all active subscribers
  var subscriberRange = sheet.getRange(startRow,startCol,numRows,numCols).getValues().filter(function(item){return item[1] == 'active'});

  // Post through Mailerlite API
  for (obj in subscriberRange){

    var field = subscriberRange[obj];

    var id = field[0];
    var firstname = field[4];
    var lastname = field[5];

    var url = "https://api.mailerlite.com/api/v2/subscribers/" + id;

    var data = {
        "fields":{
          "firstname": firstname,
          "lastname": lastname,
      }
    };
    var headers = {
              "method": "PUT",
              "contentType": "application/json",
              "headers":{"X-MailerLite-ApiDocs": "true",
                          "X-MailerLite-ApiKey": "ADD-API-KEY-HERE",
                          "User-Agent": "ReadMe-API-Explorer"},
              "payload": JSON.stringify(data)
    };

    UrlFetchApp.fetch(url, headers);
  }
}
