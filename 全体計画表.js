function txtToNo() {  //管理記号の転写
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let Mproject = sheet.getSheetByName("全体計画表");
  let lastRow = Mproject.getRange("A3").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  let startTime = new Date();

  let startNo = Number(PropertiesService.getScriptProperties().getProperty("nextNo"));
  if(!startNo) startNo = 4;  //startNoが無い場合は、4を代入。

  for (let aa = startNo; aa<=lastRow; aa++){
    let currentTime = new Date();
    let seconds = (currentTime - startTime)/1000;
    console.log(seconds);

    //300秒経過したら、トリガーを設定して中断する。
    if(seconds > 300){
      PropertiesService.getScriptProperties().setProperty('nextNo',aa);
      setTrigger();
      return;
    }

    var onlyProjectName = Mproject.getRange(aa,1).getValue();
    var Oproject = sheet.getSheetByName(onlyProjectName);
    
    var obName = Mproject.getRange(aa,2).getValue();  //品番　を取得
    var skjl = Mproject.getRange(aa,3).getValue();  //詳細　を取得
    var tanto = Mproject.getRange(aa,4).getValue();  //責任者　を取得
    var limitDay = Mproject.getRange(aa,5).getDisplayValue();  //期日日程　を取得
    var key_A = Mproject.getRange(aa,7).getValue();  //管理記号　を取得

    var lastRow2 = Oproject.getRange("A3").getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    
    for (let bb = 4; bb<=lastRow2; bb++){
      var obName2 = Oproject.getRange(bb,2).getValue();
      var skjl2 = Oproject.getRange(bb,3).getValue();
      var tanto2 = Oproject.getRange(bb,4).getValue();
      var limitDay2 = Oproject.getRange(bb,5).getDisplayValue();

      if(obName == obName2 && skjl == skjl2 && tanto == tanto2 && limitDay == limitDay2){
        Oproject.getRange(bb,7).setValue(key_A);
        console.log(key_A)
        break;
      }
    }
  }

  //処理が最後まで実行されたら、トリガーを削除する。
  let triggers = ScriptApp.getScriptTriggers();
  for(let trigger of triggers){
    if(trigger.getHandlerFunction() == "txtToNo"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  //処理が最後まで実行されたら、トリガープロパティを削除する。
  PropertiesService.getScriptProperties().deleteProperty("nextNo");

}

//トリガー設定GAS
function setTrigger(){
  let triggers = ScriptApp.getScriptTriggers();
  
  for (let trigger of triggers){
    if(trigger.getHandlerFunction() == "txtToNo"){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  //1分後にトリガーをセット
  ScriptApp.newTrigger("txtToNo").timeBased().after(1000*60).create();
  
}