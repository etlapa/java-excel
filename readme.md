### Read, Write and Update XLSX Using POI in Java
Basic application to write a new excel file, update and read it with minimum dependencies from org.apache.poi project (see pom.xml)
Just run as:

	java -jar java-excel-0.0.1-SNAPSHOT-jar-with-dependencies.jar c:/tmp/xls/myFile.xlsx

Where last parameter is the full file path location where the Excel file will be created.
It's also possible to run without this parameter:

	java -jar java-excel-0.0.1-SNAPSHOT-jar-with-dependencies.jar

in this case, the OS current temporary directory will be taken with the default file name at:

	mx.edev.java.utils.Common.FILE_NAME


This project was taken from:

	https://www.concretepage.com/apache-api/read-write-and-update-xlsx-using-poi-in-java