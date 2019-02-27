' Date Created: January 10, 2018
' Author: Louise13w
' Edited by: Ronknight
' Edit Date: 01/02/2019
' Script: search-copy-files.vbs
' Description: Copies all *-large.jpg and *.xlsx to the new location

'Path
'   images_base_path = "Z:\Images\Newest Site Product Images\whitebackground\"
	images_base_path = "original-path"
'   images_target_new_location = "E:\whitebackground\"	
	images_target_new_location = "new-path"

'Variables
   source = images_base_path
   target = images_target_new_location
   
'Objects 
   Set objFSO    = CreateObject("Scripting.FileSystemObject")
   Set objFolder = objFSO.GetFolder(images_base_path)

'Execute Main Function
	Main objFSO.GetFolder(source)
	'Prompt end of Loop
	Wscript.Echo "Copying of files completed..."
   
Sub Main(Folder)
	'List all Subfolder form Source Directory
		For Each Subfolder in Folder.SubFolders
			'Subdirectory Path (Subfolder.Path)
				Set objFolder = objFSO.GetFolder(SubFolder.Path)
				
				'Get All the files in the sub folder
					Set colFiles = objFolder.Files		
					For Each objFile in colFiles
						source_file = Subfolder.Path & "\" & objFile.Name
					
						'Search for "-large.jpg" keyword and all files that has ".xlsx" extension
						if InStr(objFile.Name, "-large.jpg") or InStr(objFile.Name, ".xlsx") then	
						'Copy to new location
							objFSO.CopyFile source_file, target, True
							Wscript.Echo "File Copied: " & objFile.Name
						End If
				Next
			
			'Recusive Iteration
				Main Subfolder
		Next
End Sub