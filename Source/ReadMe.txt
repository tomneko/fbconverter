Firebird Converter
====

Firebird Converter is a database converter from cretain version firebird to another version firebird. For example from firebird 1.5 to firebird 2.5, etc.

## Description

First we were planned to create a database converter what is from firebird 1.5 to firebird 2.5. But, if we do anyway, we decided to create the converter from firebird x.x to firebird x.x.This is done.

Next, we had to select the middleware, there were many middleware existed.We choiced ZeosDBO what support many many RDBMSs and of course firebird was supported. And of course Delphi was the only choice for build this project.We choiced Delphp XE4, because the all members had this version Delphi. :-)

We used firebird embedded library converting to local database file. Only firebird 1.0 can not treat embedded library, because the embedded library did not produce for firebird 1.0.

Of couse Firebird Converter can treat the remote database on the remote server similarly.For ecample from local file to remote server, from remote server to local file, etc.

Converting mechanism of Firebird Converter is to extract meta data from database to create database object to target database, and inserting the data from selected data. These are based on standard DDL and DML(if specified).

## Requirement

Firebird Converter can run on both Windows 32bit and 64bit.

Executable is only unpacked the archive and execute binary named "fbconverter.exe". If you are using 64bit Windows then you can use win64 package. Both win32 and win64 package include the firebird client library and embedded server what are needed versions from Firebird Converter.

## Usage

1.Execute "fbconverter.exe".
2.Specify and filled all item both Source and Destination Database.
3.You can use "Auto detect" button to select character set on Source Database.
4.You can use "Information" button to get the information about Source Database, these are the ODS Number, Page Size and Dialect.
5.You can use "Connection Test" button to test both Source and Destination Database.
6.You can specify some options on Destination Database, thease are "Create Empty Database", "Metadata Only", "Verbose mode", "Strict check", "Conversion from Integer to Bigint".
7.Then press "Convert" start the conversion.

Options:
"Create Empty Database" is to create empty database from source database information.
"Metadata Only" is to show only metadata on log window.
"Verbose mode" is to create DML at all data conversion.
"Strict check"  is to set destination firebird.conf RelaxedAliasChecking true or false. This option is valid only using embedded library.
"Conversion from Integer to Bigint"

Settings:
You can configure fbconverter.ini.

・DEFAULT_CHARSET
  is Default Character Set for Destination Database.
  Default value is SJIS_0208.
・MAX_ERROR_LINE
  is the number of lines to cut long error messages with SQL statement.
  Default value is 5.
・POST_DATA_SPLIT_SIZE
  is the number of rows for splliting process when table data is converted.
  When multi threading is enabled, one thread process the number of rows equal (POST_DATA_SPLIT_SIZE div NUMBER_OF_IMPORT_THREAD).
  Default value is 1000. 
  Minimum value is 100, if you set less than 100 then using 100.
・USE_MULTI_THREAD
  is specified whether enabled or disabled.
  Default value is enabled, this value is 1.
  If you set disabled then you set the value 0.
・NUMBER_OF_IMPORT_THREAD
  is the number of thread when multi threading is enabled.
  Default value is 4.
  Minimum value is 2, if you set less than 2 then using 2.
  Maximum value is (POST_DATA_SPLIT_SIZE div 2).
・WORK_DIR
  is the directory using with VerboseMode.
  Default value is %USERPROFILE%\AppData\Roaming\Firebird Converter

## Install

Unpack the archive where you like.

## Contribution

Programming by Hideaki Tominaga
Programming by Tsutomu Hayashi
Sponsored and Tested by Shinya Nakagawa

## Licence

[MIT](https://github.com/tomneko/fbconverter/blob/master/LICENSE)

## Author

["tomneko" Tsutomu Hayashi](https://github.com/tomneko)
["DEKO" Hideaki Tominaga](https://github.com/ht-deko)
[Shinya Nakagawa]

