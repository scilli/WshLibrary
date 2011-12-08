Wsh Library
Windows utility to delete, move and manipulate files

Copyright 2011 Simone Cilli
************************************************************
*** delete.vbs 
Delete files given a list of directory.

-	It's possible to define a list of extension (enable with enableExtCheck).
	Be carefull! If enableExtCheck = FALSE, the script delete all the files in the selected directories.
-	It's possible to define a limit age (in days) for the file (enable with enableMaxAge).
-	It's possible to iterate the operation in all the subdirectory (enable with enableSubfolders).

To use the script, open it with a text editor and set the parameter in the init section.
Than double click on the file.

Example:
folderArray = Array("C:\test\", "C:\test2\")
extensionArray = Array("log", "txt")
maxAge = 3
enableExtCheck = TRUE
enableMaxAge = TRUE
enableSubfolders = FALSE

This configuration delete all the file in the folders c:\test and c:\test2 with the extension
log or text and older than 3 days. The script doesn't delete recursevly in the subfolders.
' ************************************************************
