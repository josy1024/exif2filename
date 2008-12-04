' order nach ordnergröße umbennen.
' version: 2008_04_27
'  USAGE: cscript /nologo size2foldername.vbs foldername level
' debug = 1 (anzeigen)
' debug = 0 (umbenennen)
' Author: Ing. Josef Lahmer alias josy1024
' develcode: http://code.google.com/p/exif2filename/
' known bugs fehler mit ' im ordnername


Option Explicit
Dim refWMIService
Dim strDirName
Dim refDirectory
dim debug
dim level, worklevel
dim sh

'put the name of the directory you wish to find the size of here
strDirName="C:\daten_lokal\iso\filme"
strDirName=WScript.Arguments(0)

worklevel=WScript.Arguments(1)
debug = 0

dim asdf

if debug > 5 then WScript.echo WScript.Arguments(0) & WScript.Arguments(1)
'get reference to WMI. Also get a direct reference to the directory
'we want to know about
Set sh = Wscript.CreateObject("Wscript.Shell")
		
set refWMIService = GetObject("winMgmts:")
set refDirectory = GetObject("winMgmts:Win32_Directory='" & strDirName & "'")

'fire our recursive function and display results

	asdf = friendlybytes(getDirectorySize(refDirectory))

'clean up
set refDirectory = Nothing
set refWMIService = Nothing

'------- End of script --------

'This function returns the aggregated size
'of all files in a directory and its subdirectories
Function getDirectorySize(pCurrentDir)
	Dim numFSize
	Dim numDSize
	Dim strQuery
	Dim refItem
	Dim colFiles
	Dim colSubdirs
	numFSize = 0
	numDSize = 0
	level = level + 1

	'first get a reference to all files in the directory
	strQuery = "ASSOCIATORS OF {Win32_Directory='" & _
			wmiescape(pCurrentDir.Name) & "'} WHERE " & _
			"AssocClass=CIM_DirectoryContainsFile " & _
			"Role=GroupComponent ResultRole=PartComponent"
	set colFiles = refWMIService.ExecQuery(strQuery)

	'loop through each file and add the size of each to numFSize
	For Each refItem In colFiles
		numFSize = numFSize + refItem.FileSize
	Next
	set colFiles = Nothing

	'now get a reference to all the subdirectories
	
	strQuery = "ASSOCIATORS OF {Win32_Directory='" & _
			wmiescape(pCurrentDir.Name) & "'} WHERE " & _
			"AssocClass=Win32_SubDirectory " & _
			"ResultRole=PartComponent"
	set colSubDirs = refWMIService.ExecQuery(strQuery)

	'loop through each subdirectory, and add its
	'size to numDSize by recursively calling this function
	For Each refItem in colSubDirs
		numDSize = numDSize + getDirectorySize(refItem)
	Next
	set colSubdirs = Nothing

	'finally, print stats and return the total size
	if debug > 5 then wscript.echo worklevel & " " & level & " " & debug
	
	if (trim(worklevel) = trim(level)) then 
		rename2foldersize pCurrentDir.Name,numDSize+numFSize
	end if
	
	if ((debug > 1)) then 
		WScript.echo level & " " & pCurrentDir.Name & ": " & _
			friendlybytes(numDSize+numFSize) 
				
'		WScript.echo level & " " & pCurrentDir.Name & ": " & numFSize & _
				'" bytes in files - " & friendlybytes(numDSize) & _
				'" bytes in subdirs"
	end if
	getDirectorySize = numFSize + numDSize
	level = level - 1

End Function

function friendlybytes (size)

	if size < 1024 then
		friendlybytes = size & "b"
	elseif size < 1024*1024 then
		friendlybytes = threedigit(size,1024) & "kb"
	elseif size < 1024*1024*1024 then
		friendlybytes = threedigit(size,1024*1024) & "mb"
	elseif size < 1024*1024*1024*1024 then
		friendlybytes = threedigit(size,1024*1024*1024) & "gb"
	end if

end function

function threedigit (value, threeer)

	threedigit = fix(value/threeer*100)/100

end function

function rename2foldersize (filename, size)

	dim target, prog
	dim tsplit
	
	if (instr(1, filename, "(") > 0) then
		tsplit=split (filename, "(")
		target=tsplit(0)
	else
		target=filename
	end if
	
		prog = "%comspec% /C move """  & filename & """ """ & target & "(" & friendlybytes(size) &")" & """"
		if debug >= 0 then wscript.echo prog
		if debug = 0 then  sh.Run prog, 0, True
		
end function

function wmiescape (value)

	wmiescape=value

' not working!! 
	'	wmiescape = Replace(wmiescape, "\", "\\")

	'wmiescape = replace(wmiescape, "'", "\'") 
'	wscript.echo wmiescape

end function
