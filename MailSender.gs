function executeMailSending(){
  // ランダムで表作ってから手で変えることも考えて，再度データ読み込む
  // 最初にspreadSheetから静的に呼ぶものはここでまとめて読み込んでおく
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
      mainRange = spreadSheet.getSheetByName('Sheet1').getDataRange(),
      nameRange = spreadSheet.getSheetByName('Sheet2').getDataRange(),
      nameDict = getNameDictWithDates(getNameDict(nameRange.getValues()), mainRange.getCell(4, 4)),
      firstDate = new Date(mainRange.getCell(3, 2).getValue()),
      urlStr = spreadSheet.getUrl();
  
  var ret = Browser.msgBox("本当にメール送信しますか？", Browser.Buttons.OK_CANCEL);
  if (ret == "ok") {
    sendMails(nameDict, firstDate, urlStr)
    Browser.msgBox("メールを送信しました！");
  }else{
    Browser.msgBox("送信をキャンセルしました．");
  }
}


function sendMails(nameDict, firstDate, urlStr){
  var thisMonth = firstDate.getMonth(),
      datesStr = '',
      bodyJaText = '',
      bodyEnText = '';
  Object.keys(nameDict).forEach(function(nameKey, idx, keyArr){
    if (nameDict[nameKey].date.length !== 0){
      datesStr = nameDict[nameKey].date.map(function(date){
        return String(new Date(date.getTime() - day_offset_time_ms).getDate()); // 世界標準時とスプレッドシートの表示形式の兼ね合い(これは根本的な解決策ではない)
      }).reduce(function(prev, curr, idx, arr){
        return idx == arr.length -1 ? prev + ' and ' + curr : prev + ', ' + curr; 
      })
      bodyJaText =  "あなたの当番日は *** " + datesStr.replace(', ', '日，').replace(' and ', '日と') +"日 *** です．\n" + 
                    "\n" + 
                    "忘れないようよろしくお願いいたします．";
      bodyEnText =  "Your turn is on *** "+ datesStr + " ***" +"\n" + 
                    "\n" +
                    "Please make it sure, thank you.";      
    } else {
      bodyJaText =  "今月はあなたの掃除当番はありません．\n" + 
                    "\n" + 
                    "よろしくお願いいたします．";
      bodyEnText =  "You don't have a clean duty in this month." +"\n" + 
                    "\n" +
                    "Thank you.";    
    }
    
    if(nameDict[nameKey].mailAddress !== ''){
      MailApp.sendEmail({
        to: nameDict[nameKey].mailAddress,
        subject: (thisMonth + 1) + '月の掃除当番 (' + monthes[thisMonth] + "\'s cleanup duty)",
        body: nameDict[nameKey].name + "さん\n" + 
        "庶務係です．\n" + 
        "\n" + 
        (thisMonth + 1) + "月の掃除当番が決まりましたので連絡いたします．\n" + 
        "こちらからご確認ください．\n" + 
        urlStr + "\n" + 
        "\n" + 
        bodyJaText + "\n" + 
        "\n\n" +
        "Dear " + nameKey + "\n" + 
        "\n" + 
        monthes[thisMonth] + "'s cleanup duty was decided.\n" +
        "You can see it here: \n" + 
        urlStr + "\n" + 
        "\n" +
        bodyEnText + "\n" + 
        "\n"
      });
    }
    Utilities.sleep(1000); // 間隔空けないとうまくうごかないかも？(未検証)  
  });
}
