PROCESS_TAB = "Import";

function getZones(){
  var mulchKmlFile = DriveApp.getFileById('1Hw75YW5x5HNMlsVMj4ebpeRLfAlByHgL');
  var kmlText = mulchKmlFile.getBlob().getDataAsString('utf8');
  
  var doc = XmlService.parse(kmlText);
  var kmlns = XmlService.getNamespace('http://www.opengis.net/kml/2.2');
  var root = doc.getRootElement().getChild('Document',kmlns).getChild('Folder',kmlns);
  
  var folders = root.getChildren('Folder',kmlns);
  var zones = [];
  for (var i = 0; i < folders.length; i++) {
    var placemarks = folders[i].getChildren('Placemark',kmlns);
    for(var j=0;j<placemarks.length;j++){
      var name = placemarks[j].getChildText('name',kmlns);
      var coordsText = placemarks[j].getChild('Polygon',kmlns).getChild('outerBoundaryIs',kmlns).getChild('LinearRing',kmlns).getChildText('coordinates',kmlns);
      var coords = coordsText.split(' ').map(function(c){return c.trim().split(',').map(function(d){return parseFloat(d);}).filter(function(d){return d!=0.0;})}).filter(function(d){return d.length==2});
      var zone={name:name,
                vertx:coords.map(function(d){return d[0];}),
                verty:coords.map(function(d){return d[1];})
      };
      zones.push(zone);
    }
  }
  return zones;
}

// nvert: Number of vertices in the polygon. Whether to repeat the first vertex at the end.
// vertx, verty: Arrays containing the x- and y-coordinates of the polygon's vertices.
// testx, testy: X- and y-coordinate of the test point
function pnpoly(nvert, vertx, verty, testx, testy)
{
    var c = false;
    for (i = 0, j = nvert-1; i < nvert; j = i++) {
        if ( ((verty[i]>testy) != (verty[j]>testy)) &&
                (testx < (vertx[j]-vertx[i]) * (testy-verty[i]) / (verty[j]-verty[i]) + vertx[i]) ) {
            c = !c;
        }
    }
    return c;
}

function processZones() {
  var zones = getZones();

  var sheet = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId()).getSheetByName(PROCESS_TAB);
  var headerRow = sheet.getSheetValues(1, 1, 1, sheet.getMaxColumns());
  
  var colGeoLat = headerRow[0].indexOf('Latitude');
  var colGeoLong = headerRow[0].indexOf('Longitude');
  var colGeoErrors = headerRow[0].indexOf('Errors');
  var colZone = headerRow[0].indexOf('Zone');

  for (var i = 2; i <= sheet.getLastRow(); i++) {  
    var dataRow = sheet.getSheetValues(i, 1, 1, sheet.getMaxColumns())[0];
    if ((dataRow[colGeoErrors].toLowerCase() == 'ok') && (dataRow[colZone].trim().length < 1)) {
      var lat = dataRow[colGeoLat];
      var long = dataRow[colGeoLong];
      var matchedZone = zones.filter(d => pnpoly(d['vertx'].length, d['vertx'], d['verty'], long, lat));

      var result = 'Not found';
      if (matchedZone.length == 1) {
        result = matchedZone[0].name;
      }
      else if (matchedZone.length > 1) {
        result = 'Internal Error';
        Logger.log("Matched more than one zone for (%d, %d), this is an error", lat, long);
      }
      sheet.getRange(i, colZone + 1).setValue(result);
    }
  }
}

function extractAddress(components, target) {
  var res = components.filter(component => component.types[0] == target);
  if (res.length == 1) {
    return res[0].short_name;
  } else {
    return "-1";
  }
}

function geoCode(addr, zip) {
    var response = Maps.newGeocoder().geocode(addr);
    var status = response.status;
    var lat, long, geoStreetNum, geoStreet, geozip, errors;

    if (response.status == 'OK') {      
      var result = response.results[0];

      lat = result.geometry.location.lat;
      long = result.geometry.location.lng;
      geozip = extractAddress(result.address_components, 'postal_code');
      geoStreet = extractAddress(result.address_components, 'route');
      geoStreetNum = extractAddress(result.address_components, 'street_number');
      
      if (!geozip || geozip.trim() != zip) {
        errors = 'Zipcode Mismatch';
      } else if (geoStreetNum == "-1") {
        errors = 'Check address';
      } else {
        errors = 'OK';
      }
    }
    else {
      Logger.log('row %s %s: %s, %s', addr, result.formatted_address, result.geometry.location.lat, result.geometry.location.lng);
      errors = 'Geocoder Error'
    }
    return { 
      status: status,
      lat: lat,
      long: long,
      geoStreetNum: geoStreetNum,
      geoStreet: geoStreet,
      geozip: geozip,
      errors: errors
    };
}

function processGeoCodes() {
  var rowNum = 1;
  var sheet = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId()).getSheetByName(PROCESS_TAB);
  var headerRow = sheet.getSheetValues(rowNum, 1, 1, sheet.getMaxColumns());
  
  var colAddr = headerRow[0].indexOf('Address');
  var colZip = headerRow[0].indexOf('Zip');
  var colPickup = headerRow[0].indexOf('Pickup');
  var colGeoErrors = headerRow[0].indexOf('Errors');
  var colGeoLat = headerRow[0].indexOf('Latitude');
  var colGeoLong = headerRow[0].indexOf('Longitude');
  var colGeoAddr = headerRow[0].indexOf('GeoAddress');
  var colGeoZip = headerRow[0].indexOf('GeoZip');

  for (var i = 2; i<= sheet.getLastRow(); i++) {    
  //for (var i = 2; i<= 4; i++) {
    var dataRow = sheet.getSheetValues(i, 1, 1, sheet.getMaxColumns());

    //skip processed rows, clear the "Errors" column to re-process
    if (dataRow[0][colGeoErrors].trim().length < 1) {
      var addr = dataRow[0][colAddr] + ',Austin, TX';
      var zip  = dataRow[0][colZip].toString().trim();
      var pickup = dataRow[0][colPickup];
      //only skip geocoding if there is both no address AND Pickup is indicated
      if (zip.length < 1 && pickup == 'P') {
        // no geo decode to do
        sheet.getRange(i, colGeoErrors + 1).setValue('Self pick up');
      }
      else {
        result = geoCode(addr, zip);
          
        sheet.getRange(i, colGeoErrors + 1).setValue(result.errors);
        sheet.getRange(i, colGeoLat + 1).setValue(result.lat);
        sheet.getRange(i, colGeoLong + 1).setValue(result.long);
        sheet.getRange(i, colGeoAddr + 1).setValue(result.geoStreetNum + ' ' + result.geoStreet);
        sheet.getRange(i, colGeoZip + 1).setValue(result.geozip);      
        //Logger.log('%s: %s, %s', result.formatted_address, result.geometry.location.lat,result.geometry.location.lng);
      }
    }
  }
}

//test code
function testZone(){
  var zones = getZones();

  var found = -1;
  var lon = -97.8567071;
  var lat = 30.4281574;
  if(zones.some(function(d,i){ found = d.name; return pnpoly(d['vertx'].length,d['vertx'],d['verty'],lon,lat);})){
    var test = 'ok';
  }
  var entries = root.getChild('Folder').getChildren();
  for (var i = 0; i < entries.length; i++) {
    var title = entries[i].getChild('title', atom).getText();
    var categoryElements = entries[i].getChildren('category', atom);
    var labels = [];
    for (var j = 0; j < categoryElements.length; j++) {
      labels.push(categoryElements[j].getAttribute('term').getValue());
    }
    Logger.log('%s (%s)', title, labels.join(', '));
  }

}
