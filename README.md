# Barcode Serial Scanner

This app is a windows CLI for scanning and validating any kind of serial or psuedo unique data.

### Configuration

In this repo you will find an example config file that would be suitable for templating. The application looks in its local folder for all .ini files. If multiple are present a prompt will allow the user to select which configuration file to use. Alternatively, using the --config "filename.ini" flag as an argument at runtime allows the user to specify which file to use. In addition the user may use the --path "/location/to/file" to specify a flder to search for config files.

#### Structure
Configuration files have two sections.

1. NFO: This section specifies databse connection info and report/logging toggle.
2. dbCol: This section specifies the db column names and it's validation features.

Validation is composed as a single string with each property delimited by a colon with a single space on both sides (' : ').

##### Validation properties
* isSerial: Boolean. If set to true, the application will ensure that all newscans are unique withing the database and the current scan session.
* isMasked: Boolean. If set to true, the application will expect to find two more properties in the string, each specify a facet of the scan input masking.
* startMask: String. The application will ensure that all scans start with the specified string.
* lengthMask: Integer. The application will ensure that all scans have this many characters in total.

### Dependancies

All dependancies are availabe via pip.

* configparser
* pyodbc
* winsound
* argparse