# VHS Band Mulch Importer

This is a python script that imports a CSV export of our mulch orders from our mulch sales site (currently powered by FundTeam) into a Google Workspace spreadsheet for further processing. The target spreadsheet should have the following headers

```
Invoice	Date	Name	Email	Phone	Address	Zip	Gate Code	Validate	Pickup	Steep	Merchandise	1-Bag: Hardwood	Pallet: Hardwood	1-Bag: Black	Pallet: Black	Instructions	Latitude	Longitude	GeoAddress	GeoZip	Zone	Errors
```

The associated appscript code in `appscripts/GeoCode.js` performs geocoding/zone lookup against our KML shapes file.
There are two functions to be manually ran, `processGeocode` and `processZones`

## Running

1. Make a copy of config.yml.tmplate to config.yml
1. Edit config.yml and add appropriate code for the target spreadsheet ID. The ID is part of the spreadsheet URL, e.g. `https://docs.google.com/spreadsheets/d/<SPREADSHEET ID>/edit#gid=0`
1. `python main.py -f <csv import file>`

The first time you run this, it will ask you to go a URL to authorize the client.