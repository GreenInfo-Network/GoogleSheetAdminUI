# GoogleSheetAdminUI

An admin UI built on Google sheets: geocoding addresses, exporting as JSON, etc.

This spreadsheet and the affiliated code, allow the user to edit their Google Spreadsheet and geocode addresses, then export it as a JSON file. They can simply paste this JSON file in place of the old one, and voila -- their map is updated without using a more-complex administration UI (a separate login, a list of projects, dragging markers, etc).

The specific use case described is as follows:
* The user is accustomed to using Google Sheets. The user either provided us with access to insert these scripts, or were walked through inserting them.
* The map which is powered by this JSON file is expecting JSON but not necessarily GeoJSON. This suits a use case in which a record can have multiple points, but the 3 addresses can otherwise be treated differently (different icons for primary & secondary, for example).
* There exists some mechanism by which the user can paste the resulting JSON into the web map. This may be FTP, file upload, etc. But they would replace the map's JSON file with this new JSON file, and changes would be instantly effective.

# Credits

GreenInfo pruned and cleaned the code for our use case, but the original code was written by others.

The JSON export capability, was provided Pamela Fox http://blog.pamelafox.org/2013/06/exporting-google-spreadsheet-as-json.html

The geocoding functionality was provided by Max Vilimpoc https://github.com/nuket/google-sheets-geocoding-macro

These authors provided a valuable starting place for our needs, but the code in GoogleSheetAdminUI has been modified from their original sources. Some of the specific changes include:

* JSON and Geocode: Added help panels describing how to use them
* JSON and Geocode: Merged into one file and refactored toso additional menus-sections can be added
* Geocoding: Removal of reverse geocoding
* JSON: Skipping of blank/undefined columns
* JSON: Changed from Google UI API to Google's HtmlService API
