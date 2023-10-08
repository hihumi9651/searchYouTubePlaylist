function myFunction() {

  //URL入力ダイアログを出力
  let ytPlaylistURL = Browser.inputBox("YoutubeプレイリストのURLを入力してください");

  //URLからプレイリストのIDを抽出
  let playlistId = playlistIdExtraction(ytPlaylistURL);
  playlistId = 'PLT5klp7W4r8RYjTvxEZjOm6oqC-W6pgt1'

  //プレイリストの情報を取得し、スプレッドシートに複写
  searchYouTubePlaylistById(playlistId);

}

/**
 * YoutubeプレイリストのURLからプレイリストIDを取得する
 * @param {string} ytPlaylistURL - YoutubeプレイリストのURL
 * @return {number} playlistId - プレイリストID
 */
function playlistIdExtraction (ytPlaylistURL) {

  try {
      //playlistId以外の情報を削除
      var regexp = new RegExp(/list=/);
      var strIndexFront = ytPlaylistURL.match(regexp);

      regexp = RegExp(/&/);
      var strIndexBack = ytPlaylistURL.match(regexp);
  
      return playlistId = ytPlaylistURL.substring(strIndexFront['index'] + 5,strIndexBack['index']);

  } catch (error) {
    Logger.log('Error:', error);
  }
}

/**
 * playlistIdからプレイリストの情報を取得する
 * @param {string} playlistId - YoutubeプレイリストID
 */
function searchYouTubePlaylistById(playlistId) {

  const youtubeVideoURL = 'https://www.youtube.com/watch?v=';

  try {

    //プレイリストに含まれる動画情報を取得
    var results = YouTube.PlaylistItems.list('snippet', {
      playlistId: playlistId,
      maxResults: 1      
    });

    const mySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();

   // 取得したプレイリストデータを順次処理
   for(let i = 0; i < results.items.length; i++){

     // 追加（スプレッドシートにデータを表示）
     mySheet.getRange(i+1, 1).setValue(results.items[i].snippet.title)
     mySheet.getRange(i+1, 2).setValue(youtubeVideoURL + results.items[i].snippet.resourceId.videoId)
   }

  } catch (error) {
    Logger.log('Error:', error);
  }
}


