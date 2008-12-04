/* 
create a new album with tracks, which ar in big albums (more than 7 tracks per album)

original Script by Otto - http://ottodestruct.com  
(improved by itunes sdk)
modified by: josy1024 http://code.google.com/u/josy1024/

 */

// put your playlist name here
var hits = 7;
var 	PlaylistName = "albums: Tracks.Count > " + hits;

var	iTunesApp = WScript.CreateObject("iTunes.Application");

var	mainLibrary = iTunesApp.LibraryPlaylist;
// if you want to use some other playlist as the base to look through, uncomment this next line
// and change the "-All Songs" to the base playlist you want to use.:
//var	mainLibrary = iTunesApp.LibrarySource.Playlists.ItemByName("-All Songs");



// --- no need to change anything below this line ---
var	tracks = mainLibrary.Tracks;
var	numTracks = tracks.Count;
var 	i;
var	artistArray = new Array();
for (i = 1; i <= numTracks; i++)
{
	var	currTrack = tracks.Item(i);
	//var	artist = currTrack.Artist;
	var	artist = currTrack.Album;
	

	if ((artist != undefined) && (artist != ""))
	{
		if (artistArray[artist] == undefined)
		{
			artistArray[artist] = new Array();
		}
		artistArray[artist].push(currTrack);
	}
}
OHWPlaylist = iTunesApp.CreatePlaylist(PlaylistName);
for (var artistNameKey in artistArray)
{
	var trackArray = artistArray[artistNameKey];

	if (trackArray.length > hits)
	{
		//var	currTrack = trackArray[0];
		
		for (var trackArrayKey in trackArray)
		{
			OHWPlaylist.AddTrack(trackArray[trackArrayKey]);
		}
	}
}
