var matsunoLabCalendarID = '2sht37837mjq2hbvc8ij8fdhc0@group.calendar.google.com', //「松野研イベント予定」
    monthes = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    days = {'sun': 0, 'mon': 1, 'tue': 2, 'wed': 3, 'thu': 4, 'fri':5, 'sat':6 },
    day_offset_time_ms = 24*60*60*1000; // 世界標準時とスプレッドシートの表示形式の兼ね合い(これは根本的な解決策ではない)

// ファイルを開いたときに呼ばれる関数
function onOpen(event){
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
      menuItems = [
        {name:'シャッフル', functionName:'executeCleaningDutyPlanning'},
        {name:'カレンダー予定セット', functionName:'executeCalendarScheduling'},
        {name:'メール送信', functionName:'executeMailSending'}
      ];
    
  spreadSheet.addMenu('★', menuItems);
}

// e.g.
// nameDict = {
//     Matsuno: {name: '松野', mailAddress: 'matsuno@gmail.com', date: [担当日のDate用の数値, ...]},
//     Fukushima: {name: '福島', mailAddress: 'fukushima@gmail.com', date: [担当日のDate用の数値, ...]},
//     ...
// }
// rangeは1列目に英名，2列目に日本語名，3列目にメールアドレスがあるrange
function getNameDict(range){
  var retNameDict = {};
  range.forEach(function(arr, idx, mat){
    if(arr[0] == ''){
      range.splice(idx);
    } else {
      retNameDict[arr[0]] = {name: arr[1], mailAddress: arr[2], date: []};    
    }
  });
  return retNameDict;
}

function getNameDictWithDates(nameDict, handlingCell){
  for(;handlingCell.getValue() !== '';handlingCell = handlingCell.offset(0, 3)){
    for(var i=0, targetName = '' ; i<3 ; i++){
      targetName = handlingCell.offset(2*(i+1), 0).getValue();
      if(targetName !== ''){
        // See: makeDisplayNameStr(name)
        // ロバストにするならnameDictからマッチする名前を全探査すべきかもやけどまぁいいでしょ
        targetName = targetName.slice(targetName.indexOf('(')+1, targetName.lastIndexOf(')'));
        nameDict[targetName].date.push(handlingCell.getValue());
      }
    }
  }    
  return nameDict;
}  



 


