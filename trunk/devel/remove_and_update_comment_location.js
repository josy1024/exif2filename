/* 
    modified by: josy1024 http://code.google.com/u/josy1024/
    feature: update comment with rating and location


    ORIGINAL File: RemoveDeadTracks.js 
     
    Version: 1.0 

    Disclaimer: IMPORTANT:  This Apple software is supplied to you by Apple Computer, Inc. 
                ("Apple") in consideration of your agreement to the following terms, and your 
                use, installation, modification or redistribution of this Apple software 
                constitutes acceptance of these terms.  If you do not agree with these terms, 
                please do not use, install, modify or redistribute this Apple software. 

                In consideration of your agreement to abide by the following terms, and subject 
                to these terms, Apple grants you a personal, non-exclusive license, under AppleÕs 
                copyrights in this original Apple software (the "Apple Software"), to use, 
                reproduce, modify and redistribute the Apple Software, with or without 
                modifications, in source and/or binary forms; provided that if you redistribute 
                the Apple Software in its entirety and without modifications, you must retain 
                this notice and the following text and disclaimers in all such redistributions of 
                the Apple Software.  Neither the name, trademarks, service marks or logos of 
                Apple Computer, Inc. may be used to endorse or promote products derived from the 
                Apple Software without specific prior written permission from Apple.  Except as 
                expressly stated in this notice, no other rights or licenses, express or implied, 
                are granted by Apple herein, including but not limited to any patent rights that 
                may be infringed by your derivative works or by other works in which the Apple 
                Software may be incorporated. 

                The Apple Software is provided by Apple on an "AS IS" basis.  APPLE MAKES NO 
                WARRANTIES, EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION THE IMPLIED 
                WARRANTIES OF NON-INFRINGEMENT, MERCHANTABILITY AND FITNESS FOR A PARTICULAR 
                PURPOSE, REGARDING THE APPLE SOFTWARE OR ITS USE AND OPERATION ALONE OR IN 
                COMBINATION WITH YOUR PRODUCTS. 

                IN NO EVENT SHALL APPLE BE LIABLE FOR ANY SPECIAL, INDIRECT, INCIDENTAL OR 
                CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE 
                GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) 
                ARISING IN ANY WAY OUT OF THE USE, REPRODUCTION, MODIFICATION AND/OR DISTRIBUTION 
                OF THE APPLE SOFTWARE, HOWEVER CAUSED AND WHETHER UNDER THEORY OF CONTRACT, TORT 
                (INCLUDING NEGLIGENCE), STRICT LIABILITY OR OTHERWISE, EVEN IF APPLE HAS BEEN 
                ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 

    Copyright © 2004 Apple Computer, Inc., All Rights Reserved 
*/ 


var ITTrackKindFile	= 1;
var	iTunesApp = WScript.CreateObject("iTunes.Application");
var	deletedTracks = 0;
var	mainLibrary = iTunesApp.LibraryPlaylist;
var	tracks = mainLibrary.Tracks;
var	numTracks = tracks.Count;
var	i;

while (numTracks != 0)
{
	var	currTrack = tracks.Item(numTracks);
	
	// is this a file track?
	if (currTrack.Kind == ITTrackKindFile)
	{
		// yes, does it have an empty location?
		// (remove dead tracks)
		if (currTrack.Location == "")
		{
			// yes, delete it
			WScript.Echo (numTracks + ": DEL: " + currTrack.Artist +" - " + currTrack.Album +" - "+ currTrack.Name);
			currTrack.Delete();
			deletedTracks++;
		}
		else
		{
			//currTrack.UpdateInfoFromFile();
			if (currTrack.Podcast == true)
			{
				// podcasts do nothing..
			}
			else if (currTrack.Location.indexOf("mp2") != -1)
			{
			// mp2 tracks specals
			
				// uncomment this to update files with "mp2" found in name
				
				//currTrack.UpdateInfoFromFile();
				// currTrack.Comment = currTrack.rating + "; " + currTrack.Location + "; " + currTrack.PlayedCount;
		
				if (currTrack.Comment != currTrack.rating + "; " + currTrack.Location)
				{
					currTrack.Comment = currTrack.rating + "; " + currTrack.Location
					//WScript.Echo (numTracks + ": MP2L: " + currTrack.Location + currTrack.Artist +" - " + currTrack.Album +" - "+ currTrack.Name);
				}

			// PlayedCount (anzahl gespielt)
			// rating (wertung, 0, 20, 40, 60, 80, 100
			}
			else
			{
			// noupdate ..
			currTrack.UpdateInfoFromFile();
			//update comment with rating and location
			if (currTrack.Comment != currTrack.rating + "; " + currTrack.Location)
				{
					currTrack.Comment = currTrack.rating + "; " + currTrack.Location
					//WScript.Echo (numTracks + ": MP2L: " + currTrack.Location + currTrack.Artist +" - " + currTrack.Album +" - "+ currTrack.Name);
				}
			}
			
			
		}
		if ((numTracks % 1000) == 0)
		{
		// just for info, that script is doing something
		WScript.Echo (numTracks + ": INF: " + currTrack.Artist +" - " + currTrack.Album +" - "+ currTrack.Name);
		}
		
	}
	numTracks--;

}

if (deletedTracks > 0)
{
	if (deletedTracks == 1)
	{
		WScript.Echo("Removed 1 dead track.");
	}
	else
	{
		WScript.Echo("Removed " + deletedTracks + " dead tracks.");
	}
}
else
{
	WScript.Echo("No dead tracks were found.");
}
