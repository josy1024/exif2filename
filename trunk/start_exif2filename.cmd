@echo off
REM GET PRESET IF NOT FOUND
@SET CONFIG=exif2filename.config.cmd
if not exist %CONFIG% cp defaults\%CONFIG% .

REM CALL DEFAULT CONFIG:
REM @call defaults\%CONFIG%

REM CUSTOM CONFIGER BASEDIR
@if exist "%CONFIG%" call "%CONFIG%"

REM 2. parameter other config file:
if NOT "%1"=="" if exist "%1" call "%1"

REM echo "%1%"
REM if exist "%1%" echo "%1%"
REM if exist "%1%" call "%1%"

@echo start exif2filename
@echo DIR: %basedir%
@echo MAKEFOLDERS 0/1: %makefolders%
@if "%promptme%"=="1" pause

set para=//NOLOGO exif2filename.vbs
if exist %basedir% cscript %para% %basedir% /makefolders:%makefolders% /changefiletime:%changefiletime%
@if NOT "%basedir1%"=="" if exist %basedir1% cscript %para% %basedir1% /makefolders:%makefolders1% /changefiletime:%changefiletime1%
@if NOT "%basedir2%"=="" if exist %basedir2% cscript %para% %basedir2% /makefolders:%makefolders2% /changefiletime:%changefiletime2%
