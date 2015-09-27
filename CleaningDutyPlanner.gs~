function executeCleaningDutyPlanning(){
  // 最初にspreadSheetから静的に呼ぶものはここでまとめて読み込んでおく
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
      mainRange = spreadSheet.getSheetByName('Sheet1').getDataRange(),
      nameRange = spreadSheet.getSheetByName('Sheet2').getDataRange(),
      nameDict = getNameDict(nameRange.getValues())
      monTemplateRange = spreadSheet.getSheetByName('Sheet3').getRange(1,1,19,3),
      thuTemplateRange = spreadSheet.getSheetByName('Sheet3').getRange(1,5,19,3); 
  
  // 名前の部分をclear
  mainRange.offset(3,3,mainRange.getHeight(), mainRange.getWidth()).clear();
  // 名前を入れていく
  setCleaningDate(mainRange, monTemplateRange, thuTemplateRange, nameDict);
    
}


function setCleaningDate(mainRange, monTemplateRange, thuTemplateRange, nameDict){
  var mainRangeID = mainRange.getGridId(),
      firstDate = new Date(mainRange.getCell(3, 2).getValue()), //getCellの引数ははrangeのrelativeなOffsetで座標とは違う
      monArray = makeSameDayDatesArr(days['mon'], firstDate),
      thuArray = makeSameDayDatesArr(days['thu'], firstDate),
      cleaningDates = monArray.concat(thuArray).sort(function(a, b){return a - b;}),
      handlingCell = mainRange.getCell(4, 4),
      requiredPeopleNum = monArray.length * 2 + thuArray.length * 3,
      cleanerNameList = (function(){
        var ret = [], addArr = [], names = Object.keys(nameDict);        
        // namesは後輩から並んでいるので，必要な分だけ前部分をsliceして，それをランダムに並び替えてretにconcatする．
        // これで，後輩が当番ないのに先輩はある，ということを避ける
        Logger.log(names)
        for(;ret.length < requiredPeopleNum;){
          ret = ret.concat(names.slice(0, Math.min(names.length, requiredPeopleNum - ret.length))
                                .sort(function(){return Math.random() - 0.5;}));
        }
        return ret.sort(function(){return Math.random() - 0.5;});
      })();// 無名関数
  
  Logger.log(mainRange.getCell(4, 4).getNumberFormat())

  function makeSameDayDatesArr(targetDay, firstDate){
    var targetDates = [],
        targetDate = new Date(firstDate.getFullYear(), 
                              firstDate.getMonth(), 
                              firstDate.getDate() + (7-(firstDate.getDay()-targetDay)%7)%7);
    
    for(;targetDate.getMonth() == firstDate.getMonth();targetDate.setDate(targetDate.getDate()+7)){
      Logger.log(targetDates)
      targetDates.push(new Date(targetDate));
    }
    return targetDates;
  }
  
  function makeDisplayNameStr(name){
    return nameDict[name].name+'\n('+name+')';
  }
          
  function setNameOnCell(_date, handlingCell, nameList, nameNum, template){    
    date = new Date(_date.getTime() + day_offset_time_ms); // 世界標準時とスプレッドシートの表示形式の兼ね合い(これは根本的な解決策ではない)
    handlingCell.setValue(date);
    handlingCell.offset(0, 1).setHorizontalAlignment('right').setValue(date);
    template.copyFormatToRange(mainRangeID,
                               handlingCell.offset(0, 0).getColumn(), handlingCell.offset(1, 0).getColumn() + monTemplateRange.getWidth(),
                               handlingCell.offset(0, 0).getRow(), handlingCell.offset(1, 0).getRow() + monTemplateRange.getHeight());
    for(var i=0;i<nameNum;i++){
      handlingCell.offset(2*(i+1), 0).setValue(makeDisplayNameStr(nameList.pop()));
    }
  }
  
  cleaningDates.forEach(function(date, idx, arr){
    if(date.getDay() == days['mon']){
      setNameOnCell(date, handlingCell, cleanerNameList, 2, monTemplateRange);
    }
    if(date.getDay() == days['thu']){
      setNameOnCell(date, handlingCell, cleanerNameList, 3, thuTemplateRange);
    }
    cleanerNameList.sort(function(){return Math.random() - 0.5;});
    handlingCell = handlingCell.offset(0, 3);
  })
  
}
