// Add your own Discogs API key here:
var discogsToken = "blabla";

// Edit these two to change row background
// color according to the listing's status
var soldColor    = "#b6d7a8";
var sellingColor = "#ffffff";

// Add your USPS user here if you want to
// use that functionality
var trackingUser = "SOMEUSER";

function getRecords() {
 return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Records")
}

function getHeader() {
  var sheet = getRecords();
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
}
  
function setColor(row,color) {
  var sheet = getRecords();
  var range = sheet.getRange(row,2,1,1000);
  range.setBackground(color);
}

function makeUrl(path) {
  return "https://api.discogs.com" + path + "?token=" + discogsToken;
}

function urlFetch(url, options) {
  Utilities.sleep(Math.random()*1000);
  
  if (options) {
    return UrlFetchApp.fetch(url,options);
  } else {
    return UrlFetchApp.fetch(url);
  }
}

function fetchRelease(row) {
  var sheet = getRecords(); 
  var header = getHeader();
  var releaseIndex = header.indexOf("Release ID");
  var id = sheet.getRange(row,releaseIndex+1).getValue();
  var url = makeUrl("/releases/" + id);
  var response = urlFetch(url);
  var json = response.getContentText();
  return JSON.parse(json);
}

function setDiscogsData(sheet, row, header, label, value) {
  var index = header.indexOf(label);
  var cell = sheet.getRange(row,index+1);
  
  if (value[0] === "=") {
    sheet.getRange(row,index+1).setFormula(value);
  } else {
    sheet.getRange(row,index+1).setValue(value);
  }
}

function setThumb(sheet, row, header, data) {
  var value = "=IMAGE(\"" + data.thumb + "\",1)"
  setDiscogsData(sheet,row,header,"Thumb",value);
}

function setTitle(sheet, row, header, data) {
  setDiscogsData(sheet,row,header,"Album",data.title);
}

function setArtist(sheet, row, header, data) {
  setDiscogsData(sheet,row,header,"Artist",data.artists[0].name);
}

function setLabel(sheet, row, header, data) {
  setDiscogsData(sheet,row,header,"Label",data.labels[0].name);
}

function setLink(sheet, row, header, data) {
  var value = "=HYPERLINK(\"" + data.uri + "\",\"link\")"
  setDiscogsData(sheet,row,header,"Link", value);
}

function syncRelease(sheet, row, header) {
  var data = fetchRelease(row);
  setThumb(sheet,row,header,data);
  setArtist(sheet,row,header,data);
  setTitle(sheet,row,header,data);
  setLabel(sheet,row,header,data);
  setLink(sheet,row,header,data);
}

function getNote(sheet, header, row) {
  var listingIndex = header.indexOf("Listing ID");
  var noteCell = sheet.getRange(row, listingIndex+1);
  var data = noteCell.getNote();
  
  if (data === "") return {};
  
  return JSON.parse(data);  
}

function setNote(sheet, header, row, note) {
  var listingIndex = header.indexOf("Listing ID");
  var noteCell = sheet.getRange(row, listingIndex+1);
  noteCell.setNote(JSON.stringify(note));  
}

function onListingEdit(e) {
  var sheet = getRecords();
  var header = getHeader();
  var releaseIndex = header.indexOf("Release ID");
  var listingStatusIndex = header.indexOf("Listing Status");
  var row = e.range.getRow();
  var column = e.range.getColumn();
  var note = getNote(sheet,header,row);
  note.lastUpdated = (new Date).toLocaleString();
  setNote(sheet,header,row,note);
  
  if (column === listingStatusIndex+1) {
    if (e.value === "Sold") {
      setColor(row,soldColor);
    } else {
      setColor(row,sellingColor);
    }
   return;
  }
  
  if (column !== releaseIndex+1) return;
  if (row === "") return;
  
  syncRelease(sheet,row,header);
}

var listingHeaders = {
  release_id:       "Release ID",
  listing_id:       "Listing ID",
  status:           "Listing Status",
  condition:        "Condition",
  sleeve_condition: "Sleeve Condition",
  price:            "Price",
  comments:         "Comments",
  listed_on:        "Listed On",
  errors:           "Errors"
}

function listingData(sheet, header, index) {
  var row = sheet.getRange(index,1,1,header.length);
  var data = row.getValues()[0];
  var result = {};
  var key, index;
  
  for (key in listingHeaders) {
    index = header.indexOf(listingHeaders[key]);
    
    if (data[index]) {
      result[key] = data[index];
    }
  }
  
  return result;
}

function setListingData(sheet, header, index, data) {
  var cell, key, listingIndex;
  
  for (key in listingHeaders) {
    if (data[key]) {
      listingIndex = header.indexOf(listingHeaders[key]);
      cell = sheet.getRange(index,listingIndex+1);
      cell.setValue(data[key]);
    }
  }
}

function parseListingData(json) {
  var ret = JSON.parse(json);
  var parsed = {
    listing_id: ret.id || ret.listing_id
  }
  
  if (ret.release) parsed.release_id = ret.release.id;
  if (ret.status) parsed.status = ret.status;
  if (ret.condition) parsed.condition = ret.condition;
  if (ret.sleeve_condition) parsed.sleeve_condition = ret.sleeve_condition;
  if (ret.price) parsed.price = ret.price.value;
  if (ret.comments) parsed.comments = ret.comments;
  if (ret.listed_on) parsed.listed_on = ret.listed_on;
  
  return parsed;
}

function fetchDiscogsListing(sheet, header, index) {
  var data = listingData(sheet, header, index);
  var url, response, ret, content;
  
  if (!data.listing_id) return;
  
  url = makeUrl("/marketplace/listings/" + data.listing_id);
  response = urlFetch(url);
  ret = parseListingData(response.getContentText());
  
  // Only update when ret status is more advanced.
  if (ret.status !== "Sold") return;

  setListingData(sheet, header, index, ret);
}

function syncDiscogsListing(sheet, header, index) {
  var data = listingData(sheet, header, index);
  var note = getNote(sheet,header,index);
  var options, errors, response, ret;
  
  // First make a get on existing listing if For Sale to see if
  // it has been sold.
  if (data.listing_id && data.status == "For Sale") {
    fetchDiscogsListing(sheet, header, index);
    data = listingData(sheet, header, index);
  }

  if (data.status == "Sold") return;
  
  // Stop here if last updated <= last submitted
  if (!note.lastUpdated && note.lastSubmitted) return;
  
  if (note.lastUpdated && note.lastSubmitted &&
      ((new Date(note.lastUpdated)) <= (new Date(note.lastSubmitted)))) return;
  
  errors = [];
  
  var checkError = function(key,values) {
    if (values.indexOf(data[key]) == -1) {
      var label = listingHeaders[key];
      errors.push("Error: " + label + " should be one of: " + values.join(", "));
    }
  }
  
  var validStates = ["Draft", "For Sale"];
  checkError("status",validStates);
  
  var validConditions = [
    "Mint (M)", "Near Mint (NM or M-)",
    "Very Good Plus (VG+)", "Very Good (VG)",
    "Good Plus (G+)", "Good (G)",
    "Fair (F)", "Poor (P)"
  ];
  checkError("condition",validConditions);
  
  var validSleeveConditions = [
    "Mint (M)", "Near Mint (NM or M-)",
    "Very Good Plus (VG+)", "Very Good (VG)",
    "Good Plus (G+)", "Good (G)",
    "Fair (F)", "Poor (P)",
    "Generic", "Not Graded", "No Cover"
  ];
  checkError("sleeve_condition",validSleeveConditions);
  
  if (errors.length > 0) {
    setListingData(sheet, header, index, {errors: errors.join(", ")});
    return
  }
  
  if (data.listing_id) {
    url = makeUrl("/marketplace/listings/" + data.listing_id);
  } else {
    url = makeUrl("/marketplace/listings");
  }
  
  options = {
    "method":      "post",
    "contentType": "application/json"
  };
  
  options.payload = JSON.stringify(data);
  response = urlFetch(url,options);

  note.lastSubmited = (new Date).toLocaleString();
  setNote(sheet,header,index,note);
  
  if (!data.listing_id) {
    ret = parseListingData(response.getContentText());
    setListingData(sheet, header, index, ret);
  }
}
  
function syncDiscogsListings() {
  var sheet = getRecords();
  var lastRow = sheet.getLastRow();
  var header = getHeader();
  
  var index;
  for (index = 2; index <= lastRow; index++) {
    syncDiscogsListing(sheet,header,index);
  }
}

function updateTracking(sheet, header, row) {
  var trackingIDIndex = header.indexOf("Tracking ID");
  var lastTrackingIndex = header.indexOf("Last Tracking Status");
  var id = "" + sheet.getRange(row,trackingIDIndex+1).getValue();
  
  if (id === "") return;
  
  var baseUrl = "http://production.shippingapis.com/ShippingAPI.dll?API=TrackV2&XML=";
  var xml = "<TrackRequest USERID=\"" + trackingUser + "\">\
               <TrackID ID=\"" + id + "\"></TrackID>\
             </TrackRequest>";
  var url = encodeURI(baseUrl + xml);
  var xml = UrlFetchApp.fetch(url).getContentText();
  var document = XmlService.parse(xml);
  var root = document.getRootElement();
  var info = root.getChildren()[0];
  var detail = info.getChildren()[0];
  
  sheet.getRange(row,lastTrackingIndex+1).setValue(detail.getText());
}

function syncTrackings() {
  var sheet = getRecords();
  var lastRow = sheet.getLastRow();
  var header = getHeader();
  
  var index;
  for (index = 2; index <= lastRow; index++) {
    updateTracking(sheet,header,index);
  }
}
