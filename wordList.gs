function main() {
  // スプレッドシートを開く
  var spreadsheet = SpreadsheetApp.openById('12OKvcnrQw6IAj_lkw_1L51agY6K6JPqg0PHiN-wOLE8');
  var sheetWordList = spreadsheet.getSheetByName('単語リスト');
  var sheetResearch = spreadsheet.getSheetByName('再検索');
  
  // 既存の単語リストを取得
  const wordListTemp = sheetWordList.getLastRow() == 1 ? [] : sheetWordList.getRange(2, 1, sheetWordList.getLastRow() - 1).getValues();
  var wordList = []; // すでに追加されている単語のリスト
  for(var num in wordListTemp){
    wordList.push(wordListTemp[num][0]);
  }
  

  //GMailからデータを取得
  var texts = getTextFromGmail('メモのエクスポート label:inbox');
  
  // 手に入れたテキストについて、
  for(numText in texts){
    var items = texts[numText].match(/<div class="noteText">[\s\S]*?<\/div>/gi);
    for (num in items){
      var word = items[num].match(/>([\s\S]*?)</i)[1].replace(/[^a-z]*/g, '');
      if(wordList.indexOf(word) != -1){
        continue;
      }
      dic = searchWord(word);
      if(dic['word'] == null){
        sheetResearch.appendRow([word]);
      }else if(wordList.indexOf(dic['word']) == -1 && dic['meaning'] !== ''){
        sheetWordList.appendRow([dic['word'],dic['meaning'],dic['level']]);
        wordList.push(dic['word']);
      } 
    }
  }
  
  // 再検索の単語リストを取得
  var researchWordList = sheetResearch.getLastRow() == 0 ? [] : sheetResearch.getRange(1, 1, sheetResearch.getLastRow()).getValues(); // 検索ミスした単語のリスト
  
  // 再検索の単語について逆順に意味を調べる
  var dic;
  const inverse = researchWordList.length - 1;
  for (var num in researchWordList){
    dic = searchWord(researchWordList[inverse - num][0]);
    if(dic['word'] != null){
      if(dic['meaning'] == ''){
      }else if(wordList.indexOf(dic['word']) == -1){
        sheetWordList.appendRow([dic['word'],dic['meaning'],dic['level']]);
      }
      sheetResearch.deleteRow(inverse - num + 1);
    }
  }

}


// クエリによる検索結果の最初のメッセージの最初の添付ファイルのテキストを取得
// メールはアーカイブフォルダに移動される
function getTextFromGmail(query){
  var textList = [];
  var threads = GmailApp.search('メモのエクスポート label:inbox');
  for(var numThread in threads){
    textList.push(threads[numThread].getMessages()[0].getAttachments()[0].getDataAsString());
    //threads[numThread].moveToArchive();
  }
  
  return textList;
}

// 単語をweblioで検索
function searchWord(word){
  Logger.log('search: ' + word);
  var dic = Object.create(null);
  
  const url_base = 'http://ejje.weblio.jp/content/';
  try {
    const html = UrlFetchApp.fetch(url_base + word).getContentText();
  } catch(e){
    Logger.log('word:' + word + ', error:' + e.message);
    return dic
  }
  
  var meaning = html.match(/<td\ class=content-explanation>([\s\S]*?)<\/td>/i);
  
  // 意味が見つからなかった場合
  if (meaning == null){
    dic['word'] = word;
    dic['meaning'] = '';
    dic['level'] = '99';
    return dic
  }
  
  meaning = meaning[1].replace(/<b>|<\/b>/gi,'');
  const level_search = html.match(/<span\ class=\"learning-level-label\">レベル<\/span>[\s\S]*?<span\ class=[\"|]learning-level-content[\"|]>([\s\S]*?)<\/span>/i);
  var level;
  if(level_search == null){
    level = '31';
  }else{
    level = level_search[1]; 
  }
  dic['word'] = word;
  dic['meaning'] = meaning;
  dic['level'] = level;
  
  //三単現、活用、複数形への対処
  reword = meaning.match(/([a-zA-Z]*?)[\s]*の(三人称単数現在|複数形|過去形|過去分詞|現在分詞)/i);
  if(reword != null){
    dic = searchWord(reword[1]);
  };
  return dic
}

