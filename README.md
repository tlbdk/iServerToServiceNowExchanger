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

Show help
--help

Download to file. Either type servicenow OR iserver after the -d.
If ServiceNow is chosen, it will download the RelationsTable and the ObjectsTable by using the <to file> as directory/prefix. Afterwards, it merges the 2 tables into 1 and saves it as <to file>.
If iServer is chosen, it will just download the file to <to file>.

-d <servicenow/iserver> -f <to file>

Upload file. Either type servicenow OR iserver after the -u.
If ServiceNow is chosen, it will upload the RelationsTable and ObjectsTable by using the <to file> as directory/prefix.
if iServer is chosen, it will just upload the file.

-u <servicenow/iserver> -f <from file>

Merge files. Manually merge multiple files by listing them as: -m "1" -m "2" ... -m "n" -f <to file>. The sheets will be in the same order as the files.

-m <file_1> -m <file_n> -f <to fileMerged>

Split 1 or more files into 1 pr worksheet (adds worksheet name as postfix automatically).

-s <file_1> -s <file_n>