' bildersortierscript: makedirs_of_jpg_renametime.vbs
' Author: Ing. Josef Lahmer alias josy1024
' version: 2008_04_25
' http://code.google.com/p/exif2filename/


public debug 
public msg, filetimetoken(6)
public makefolders
public changefiletime
public strSourcefolderpath

CONST SPACER = "_"
CONST SPACERBETWEEN = "-"

' makefolder paramter 0 or 1 =>
'make folders with dayname
makefolders = 0
changefiletime = 0

debug = 1

If WScript.Arguments.Named.Exists("debug") Then
	debug = WScript.Arguments.Named("debug")
End If

If WScript.Arguments.Named.Exists("makefolders") Then
	makefolders = WScript.Arguments.Named("makefolders")
End If

If WScript.Arguments.Named.Exists("changefiletime") Then
	changefiletime = WScript.Arguments.Named("changefiletime")
End If

strSourcefolderpath = verifypath(wscript.arguments(0))
	
msg = "Ordner: " & strSourcefolderpath & vbCRLF & vbCRLF
	
foldersorter strSourcefolderpath

wscript.echo msg

function verifypath (folder)

	if (Right(folder,1) <> "\") then
		verifypath = folder & "\"
	else
		verifypath = folder
	end if
end function

function foldersorter (strSourcefolderpath)

	Dim aFolderArraySource
	Dim aFolderArrayDestination
	Dim FolderListSource
	Dim FolderListDestination
	Dim oFolderSource
	Dim oFolderDestination
	Dim bSourceExists
	Dim bDestinationExists

	'On Error Resume Next

	Dim oFSO

	Set oFSO = CreateObject("Scripting.FileSystemObject")

	Set aFolderArraySource = oFSO.GetFolder(strSourcefolderpath)
	' Set aFolderArrayDestination = oFSO.GetFolder(strDestinationfolderpath)
	Set FolderListSource = aFolderArraySource.SubFolders
	' Set FolderListDestination = aFolderArrayDestination.SubFolders

	'1. foreach file in folder, 
	'get filemodifiedtime,
	'if ! folderexist createfolder
	'movefile2folder


	Set aFileArraySource = oFSO.GetFolder(strSourcefolderpath)
	Set FileListSource = aFileArraySource.Files

	For each oFileSource in FileListSource
	  	
	  	if debug > 2 Then wscript.echo "For each oFileSource: " & oFileSource.Name
	  	
'	  	If (Len (oFileSource.name) > 4) AND ((UCASE(Right(oFileSource.Name,3)) = "JPG") OR (Right(oFileSource.Name,3) = "AVI"))  Then
	  	If (Len (oFileSource.name) > 4) AND ((UCASE(Right(oFileSource.Name,3)) = "JPG") OR (UCASE(Right(oFileSource.Name,3)) = "JPE") ) Then
	  		
			renameto = renamer(oFileSource.Name)

			if makefolders = 1 then
				targetpath = filetimetoken(1) & SPACER & filetimetoken(2) & SPACER & filetimetoken(3)
				
				if debug > 2 Then wscript.echo "such target: " & strSourcefolderpath & targetpath
				
				targetpath = suchfolder2(strSourcefolderpath & targetpath)
				
				if debug > 2 Then wscript.echo "create target: " & targetpath
				if Not oFSO.FolderExists(targetpath) Then oFSO.CreateFolder (targetpath)
				renameto = targetpath & "\" & renameto
			else
				' 2007_02_25-15_03_23-cimg2839745.jpg
				renameto = strSourcefolderpath  & renameto
			end if
		  	' oFileSource.Copy strDestinationfolderpath & "\" & oFileSource.Name
			
			
			msg = msg & renameto & vbCRLF

			' if debug > 2 Then wscript.echo renameto
			if debug > 0 Then wscript.echo renameto
			on error resume next
		  	oFileSource.Move renameto
			
		   if err.number = 58 then
		   ' datei bereits vorhanden error exception
				oFileSource.Delete
			elseif err.number <> 0 then 
				msgbox  err.number  &  err.description & renameto
				err.clear
			end if
		End If
	  Next

end function

function cleanupfilename(filename)
	dim tokens
'		2007_01_19-22_04_04-IMG_0811.JPG

	tokens = Split(filename, SPACERBETWEEN)
	
	if debug > 2 Then wscript.echo "cleanupfilename" & filename & " " & UBound(tokens) & " " & LBound(tokens)
	if (UBound(tokens) >= 2) then
		cleanupfilename = trim(tokens(UBound(tokens)))
	else
		cleanupfilename = filename
	end if	
end function

function addnulls(val)
	
	if (val < 10) then
		addnulls = "0" & cstr(val)
	else
		addnulls = val
	end if
	
end function 

function renamer (filename)
	dim tokens, tokens2
	dim newtime
	
	'oFileSource.Name
	exifinfo = getexif (filename, "Image timestamp")
	'DTOrig : 2007:01:19 12:18:04
	'wscript.echo "in renamer:" & filename & ": " & exifinfo
	tokens = Split(exifinfo, ":")
	tokens2 = Split(tokens(3), " ")
	
	filetimetoken(0) = cleanupfilename(filename)
	filetimetoken(1) = trim(tokens(1))
	filetimetoken(2) = trim(tokens(2))
	filetimetoken(3) = trim(tokens2(0))
	filetimetoken(4) = trim(tokens2(1))
	filetimetoken(5) = trim(tokens(4))
	filetimetoken(6) = trim(tokens(5))


	if ( changefiletime = 1 ) then
		newtime = correcttime (filetimetoken(0), filetimetoken(1),filetimetoken(2),filetimetoken(3),filetimetoken(4),filetimetoken(5),filetimetoken(6))
		filetimetoken(1) = year (newtime)
		filetimetoken(2) = addnulls(month (newtime))
		filetimetoken(3) = addnulls(day (newtime))
		filetimetoken(4) = addnulls(hour (newtime))
		filetimetoken(5) = addnulls(minute (newtime))
		filetimetoken(6) = addnulls(second (newtime))
	end if
	
	
	renamer = filetimetoken(1) & SPACER & filetimetoken(2) & SPACER & filetimetoken(3) & _
		SPACERBETWEEN & filetimetoken(4) & SPACER & filetimetoken(5) & SPACER & filetimetoken(6) & _
		SPACERBETWEEN & filetimetoken(0)
	
end function


function correcttime (filename, y, m, d, h, n, s)
' http://www.vbarchiv.net/commands/cmd_dateadd.html
	dim checkfile
	dim datevar
	
	checkfile = lcase(filename)
	
		
	if (instr(checkfile, "xxxxxxcimg") > 0) then
		datevar = DateAdd("h", -2, d & "." & m & "." & y & " " & h & ":" & n & ":" & s)
	elseif (instr(checkfile, "pict") > 0) then
	'h=stunde
	'n=minute
		datevar = DateAdd("h", 0, d & "." & m & "." & y & " " & h & ":" & n & ":" & s)
		datevar = DateAdd("d", +964, datevar)
		datevar = DateAdd("h", +4, datevar)
		datevar = DateAdd("n", +48, datevar)
		wscript.echo "pict: " &  datevar
	else
		datevar = DateAdd("h", 0, d & "." & m & "." & y & " " & h & ":" & n & ":" & s)
	end if
	
	correcttime =  datevar	
	
end function

Function getexif (filename, key)

	Set sh = Wscript.CreateObject("Wscript.Shell")
	Set env = sh.Environment("PROCESS")
		
	Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
	Set objShell = Wscript.CreateObject("Wscript.Shell")
	rundirname = GET_SCRIPT_Verzeichnis
	
	objName = objFSO.GetTempName
	objTempFile = env("tmp") & "\" & objName
	prog = "%comspec% /C "" """ & rundirname & "exiv2.exe"" """ & strSourcefolderpath & filename & """ | " & _
 		"find /I """ & key & """ >" & objTempFile & """"

		wscript.echo prog
	
	sh.Run prog, 0, True
	
	

	Set objTextFile = objFSO.OpenTextFile(objTempFile, 1)
	Do While objTextFile.AtEndOfStream <> True
		strText = objTextFile.ReadLine
	Loop

	objTextFile.Close
	objFSO.DeleteFile(objTempFile)
	
	getexif = strText
end function

Function FileLastModified(Fname)

  FileLastModified = ""

  Set fs = CreateObject("Scripting.FileSystemObject")

  if fs.FileExists(Fname) = True then
    Set f = fs.GetFile(Fname)
    FileLastModified = f.DateLastModified
  end if
Set f = Nothing
Set fs = Nothing
End Function

Function GET_SCRIPT_Verzeichnis()
	Dim strPfad
	Dim intLaenge
	strPfad=WScript.ScriptFullName
	intLaenge=Len(WScript.ScriptName)
	strPfad=Mid(strPfad,1,Len(strPfad)-intLaenge)
	
	'If strPfad = "" Then
	'		strPfad="\\server-gpm\autoinstall\"
	'End If

	GET_SCRIPT_Verzeichnis=strPfad
End Function


function suchfolder2(objPath)
	
	'objPath = Replace(pathname, "\", "\\")
	
	if debug > 2 Then 	wscript.echo "suchfolder2: " & objPath
	dir = dirname(objPath)
	base = basename(objPath)
	
	suchfolder2 = objPath
	
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.Namespace(dir)
	Set objFolderItem = objFolder.Self
	
	Set colItems = objFolder.Items
	For Each objItem in colItems
		'if debug > 2 Then 	wscript.echo "instr" & instr(objItem.Name, ".") & len(objItem.Name) 
		' wenn keine dateierweiterung
		if (instr(objItem.Name, ".") <> (len(objItem.Name) - 3 )) then
			if (left(objItem.Name, len(base)) = base) then
				suchfolder2 = dir & "\" & objItem.Name 
			end if
		end if
	Next
end function

function dirname(fullname)

	dim c1
	dim c2
	
	C2 = InStr(C1 + 1, fullname, "\")
	While C2 > 0
	   'dirname = Mid(fullname, C1 + 1, C2 - C1 - 1)
	   
	   C1 = C2
	   C2 = InStr(C1 + 1, fullname, "\")
	Wend
	    
	dirname= Left(fullname, C1 - 1)
end function

function basename(fullname)

	C2 = InStr(C1 + 1, fullname, "\")
	While C2 > 0
	   basename = Mid(fullname, C1 + 1, C2 - C1 - 1)
	   
	   C1 = C2
	   C2 = InStr(C1 + 1, fullname, "\")
	Wend
	    
	basename = Right(fullname, Len(fullname) - C1)
end function

