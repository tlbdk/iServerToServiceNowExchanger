How to use iServerToServiceNowExchanger (Working title!)
========================================================

The application has no interface and is used in Command Prompt be changing values in the App.config and adding parameters when started.

App.config
----------
The config has a number of settings, most of them self-explanatory.
The only thing to remember is to use the escape characters below when typing the URL.

	"   &quot;
	'   &apos;
	<   &lt;
	>   &gt;
	&   &amp;


Parameters
----------
The parameters is typed without the "<" and ">" characters.
When typing ServiceSow and iServer, it is case-insensitve.
When naming the file, the extention will be stripped and be replaced by .xls as this is the only supported filetype.

__Show help__

	--help

__Download to file__
Either type servicenow OR iserver after the -d.
If ServiceNow is chosen, it will download the RelationsTable and the ObjectsTable by using the "to file" as directory/prefix. Afterwards, it merges the 2 tables into 1 and saves it as "to file".
If iServer is chosen, it will just download the file to "to file".

	-d <servicenow/iserver> -f <to file>

__Upload file__
Either type servicenow OR iserver after the -u.
If ServiceNow is chosen, it will upload the RelationsTable and ObjectsTable by using the "to file" as directory/prefix.
if iServer is chosen, it will just upload the file.

	-u <servicenow/iserver> -f <from file>

__Merge files__
Manually merge multiple files by listing them as: -m "1" -m "2" ... -m "n" -f "to file". The sheets will be in the same order as the files.

	-m <file_1> -m <file_n> -f <to fileMerged>

__Split__
Split 1 or more files into 1 pr worksheet (adds worksheet name as postfix automatically).

	-s <file_1> -s <file_n>


Examples
--------

###Case 1
To download a file, run:

	iServerToServiceNowExchanger.exe -d SERVICENOW -f "C:\some\dir\foo.bar"

The result will be 3 files in the C:\some\dir\ directory:

	foo_ObjectsTable.xls
	foo_RelationsTable.xls
	foo_Merged.xls

The 2 first are the downloaded files and the last one is the 2 files merged.

###Case 2
Continuing Case 1, these file could be uploaded.

	-u ServiceNow -f "C:\some\dir\foo.bar"

This will upload the 2 files:

	foo_ObjectsTable.xls
	foo_RelationsTable.xls

To the upload URLs defined in App.config:

	<?xml version="1.0" encoding="utf-8" ?>
	<configuration>
	    <appSettings> 
	      ...
	      <add key="serviceNowUploadURLObject" value="YourWebsite.com/uploadPlaceForObjects" />
	      <add key="serviceNowUploadURLRelation" value="YourWebsite.com/uploadPlaceForRelations" />
	      ...
	    </appSettings>
	</configuration>