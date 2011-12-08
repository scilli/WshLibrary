' ************************************************************
' DELETE
' This script delete files with specific attributes
' (path, extension, age)
'
' Copyright 2011 Simone Cilli
' ************************************************************
' INIT:
' List of folders
folderArray = Array("C:\test\")
' List of extensions
extensionArray = Array("log")
' Max age of files (in days)
maxAge = 1
' Flag
enableExtCheck = TRUE ' Check Extensions
enableMaxAge = FALSE ' Check file's max age
enableSubfolders = FALSE ' Delete files also in the subdirectories
' ************************************************************
' CODE:
' Main
SET objFso = CREATEOBJECT("Scripting.FileSystemObject")
Delete folderArray, extensionArray, maxAge, enableExtCheck, enableMaxAge, enableSubfolders
wscript.echo "Operation completed successfully"
' Function
SUB Delete(BYVAL folderArray, BYVAL extensionArray, BYVAL maxAge, BYVAL enableExtCheck, BYVAL enableMaxAge, BYVAL enableSubfolders)
	FOR EACH folder in folderArray ' Iterate on folders
		IF  NOT objFso.FolderExists(folder) THEN ' Check if the folder exists
			wscript.echo "Folder " & folder & " doesn't exist"
		ELSE
			'wscript.echo "Clean " & folder
			SET objFolder = objFso.GetFolder(folder)
			DeleteFiles objFolder, extensionArray, maxAge, enableExtCheck, enableMaxAge, enableSubfolders
		END IF
	NEXT
END SUB
SUB DeleteFiles(BYVAL objFolder, BYVAL extensionArray, BYVAL maxAge, BYVAL enableExtCheck, BYVAL enableMaxAge, BYVAL enableSubfolders)
	FOR EACH objFile IN objFolder.Files ' Iterate on files
		'wscript.echo "File " & objFile.path
		FOR EACH ext IN extensionArray
			IF NOT enableExtCheck OR RIGHT(UCASE(objFile.path),LEN(ext)+1) = "." & UCASE(ext) THEN ' Check extension
				IF NOT enableMaxAge OR objFile.DateLastModified < (NOW - maxAge) THEN' Check age
					'wscript.echo "Delete file " & objFile.path
					objFile.Delete ' Delete file
					EXIT FOR
				END IF
			END IF
		NEXT
	NEXT
	IF enableSubfolders = TRUE THEN ' Recursive call of DeleteFiles function for all the subdirectories
		FOR EACH objSubfolder in objFolder.SubFolders
			DeleteFiles objSubfolder, extensionArray, maxAge, enableExtCheck, enableMaxAge, enableSubfolders
		NEXT
	END IF
END SUB