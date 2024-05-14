function addVideos() {
  const sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // regular expression used to parse the various formats of Youtube video URLs
  const regExp = /(?:(?:\?v=)|(?:youtu\.be\/)|(?:youtube\.com\/embed\/))([^\&\n\?]+)/i;

  // asks for the ID of the Youtube playlist to add videos to
  var playlist_prompt = ui.prompt("Enter playlist id (list parameter in url, for example: youtube.com/watch?v=2...b0&list=PL...7X is PL...7X)");
  const playlist_id = playlist_prompt.getResponseText();

  // asks for the column number where the URLs are located
  var column_prompt = ui.prompt("Enter column number where urls are present (A=1, B=2, C=3,...)(also assumes there is a header in the first row, starts in the second)");
  var column = parseInt(column_prompt.getResponseText());

  // grabs the URLs from the sheet, starting in the second row and collecting all the elements in the column supplied in the above prompt
  const video_urls = sheet.getRange(2, column, sheet.getLastRow()).getValues().flat([1]);

  // arrays for splitting apart which URLs are valid Youtube video links and which aren't
  const parseable_urls = [];
  const unparseable_urls = [];
  
  // iterates over all collected elements and splits them between the above arrays based on whether the regex matches or not
  video_urls.forEach(url => {
    var id = regExp.exec(url);
    if (id == null) {
      unparseable_urls.push(url);
    } else {
      parseable_urls.push(url);
    }
  });

  // array for the URLs that Youtube returns a "Video not found." error for, indicating the video was removed for some reason
  const unavailable_video_urls = [];

  // iterates through parseable URLs, matching each to extract the video ID and using that to insert to the playlist via the Youtube API
  parseable_urls.forEach(url => {
    var matched_regex = regExp.exec(url);
    var video_id = matched_regex[1];
    try {
      YouTube.PlaylistItems.insert({

        snippet: {
                
          playlistId: playlist_id,
                
          resourceId: {
                    
            kind: "youtube#video",
                    
            videoId: video_id
                    
            }
          }
        }, "snippet");
    } catch (e) {
      // if there is an error of the video not being found, then that URL is added to an array
      if (e["details"]["message"] == "Video not found.") {
        unavailable_video_urls.push(url);
        // print("Video not found.");
      }
    }
  });

  // add a label to each array that will show up when appending these to the bottom of the sheet
  unparseable_urls.unshift("Unparseable URLs:");
  unavailable_video_urls.unshift("Unavailable Video URLs:");

  // append the unavailable and unparseable URLs to the bottom of the sheet 
  sheet.appendRow(unparseable_urls);
  sheet.appendRow(unavailable_video_urls);
}

function onOpen() {
  // add UI element to call the above function
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('YouTube Data')
  .addItem('Add videos to playlist', 'addVideos')
  .addToUi();
}
