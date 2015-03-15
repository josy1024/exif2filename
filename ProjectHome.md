vba windows scripts that read out exif file infos pics and change the filename to the picture snapshot time.

# functions #
**time chronological order of images** different cam types
**modify timestamp**

```
Input file: CIMG1058.JPG
Input Image timestamp : 2007:09:05 23:12:17
Output file: 2007_09_05-23_12_17-CIMG1058.JPG
```

## setup ##
**prerequisits:http://www.exiv2.org
extract exiv2.exe tool into the same folder like the vbs and cmd-scripts.**

**download http://exif2filename.googlecode.com/svn/trunk/exif2filename.vbs**


## usage ##
```
cscript exif2filename.vbs /makefolders:1 /debug:1
cscript exif2filename.vbs /makefolders:0 /debug:1
cscript exif2filename.vbs /makefolders:1 /debug:1 /changefiletime:1 (see function correcttime)
```
## changelog ##

## todo ##
**documentation** updater über wget?
```
wget -c http://exif2filename.googlecode.com/svn/trunk/exif2filename.vbs
```
