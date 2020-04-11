//Project COVID-19 PULL DATA From API
//Connect with google app script
//connect google sheet

var lineClient = new LineBotSDK.Client({
    channelAccessToken: 'Your token channel access token.',
  });
  function doPost(e) {
    var ssId = "Your google sheet id";
    var ss = SpreadsheetApp.openById(ssId);
    var sheetUsers = ss.getSheetByName("users");
  
    //use BetterLog
    Logger = BetterLog.useSpreadsheet(ssId);
  
    //Logger.log("Hello from BetterLog :)");
  
    var requestJSON = e.postData.contents;
    //Logger.log(requestJSON);
    
    var requestObj = JSON.parse(requestJSON).events[0];
    
    var token = requestObj.replyToken;
    
    if (requestObj.type === "follow") {
      var userId = requestObj.source.userId;
      Logger.log("This is user Id: " + userId);
      
      var userProfiles = getUserProfiles(userId);
      
      var lastRow = sheetUsers.getLastRow();
      sheetUsers.getRange(lastRow + 1, 1).setValue(userId);
      sheetUsers.getRange(lastRow + 1, 2).setValue(userProfiles[0]);
      sheetUsers.getRange(lastRow + 1, 3).setValue(userProfiles[1]);
      sheetUsers.getRange(lastRow + 1, 4).setFormula("=image(C" + (lastRow + 1) + ")");
      
      var replyText = getCovidSheet();
      return replyMessageNewUser(token, replyText);
    }
  
    var userMessage = requestObj.message.text;
    //Logger.log(userMessage);
    //var replyText = JSON.stringify(requestObj);
    if(userMessage.includes("ทั้งหมด") || userMessage.includes("ล่าสุด") || userMessage.includes("รายวัน") || userMessage.includes("ยอดรวม") ){
      var replyTextFromSheet = getCovidSheet();
      return replyMessage(token, replyTextFromSheet);
    }
    var replyText = getByProvince(userMessage)
    return replyMessage(token, replyText);
  }
  
  function getUserProfiles(userId) {
    var url = "https://api.line.me/v2/bot/profile/" + userId;
    var lineHeader = {
      "Content-Type": "application/json",
      "Authorization": "Bearer [Your token channel access token.]"
    };
    
    var options = {
      "method" : "GET",
      "headers" : lineHeader
    };
    
    var responseJson = UrlFetchApp.fetch(url, options);
    
    Logger.log("User Profiles Response: " + responseJson);
    
    var displayName = JSON.parse(responseJson).displayName;
    var pictureUrl = JSON.parse(responseJson).pictureUrl;
    
    return [displayName, pictureUrl];
  }
  function replyMessage(token, replyText) {
    var url = "https://api.line.me/v2/bot/message/reply";
    var lineHeader = {
      "Content-Type": "application/json",
      "Authorization": "Bearer [Your token channel access token.]"
    };
  
    var postData = {
      "replyToken" : token,
      "messages" : [{
        "type" : "text",
        "text" : replyText
      }]
    };
  
    var options = {
      "method" : "POST",
      "headers" : lineHeader,
      "payload" : JSON.stringify(postData)
    };
  
    try {
      var response = UrlFetchApp.fetch(url, options);
  //    var todayCovidtext = getCovidToday();
  //    if(flag==true){
  //      lineClient.replyMessage(token, { type: 'text', messages: todayCovidtext });
  //    }
    }
    
    catch (error) {
      Logger.log(error.name + "：" + error.message);
      return;
    }
      
    if (response.getResponseCode() === 200) {
      //Logger.log("Sending message completed.");
    }
  }
  
  function replyMessageNewUser(token, replyText) {
    var url = "https://api.line.me/v2/bot/message/reply";
    var lineHeader = {
      "Content-Type": "application/json",
      "Authorization": "Bearer [Your token channel access token.]"
    };
  
    var postData = {
      "replyToken" : token,
      "messages" : [
        {
        "type" : "text",
        "text" : replyText
        },
        {
        "type" : "text",
        "text" : "************ วิธีใช้งาน ************\nสามารถเลือกแสดงตาม \"จังหวัด\" โดยพิมพ์ข้อมูล เช่น \"กรุงเทพมหานคร\" \nหรือต้องการเรียกดูทั้งหมด ให้พิมพ์ \"ล่าสุด\" หรือ \"ทั้งหมด\""
      }]
    };
  
    var options = {
      "method" : "POST",
      "headers" : lineHeader,
      "payload" : JSON.stringify(postData)
    };
  
    try {
      var response = UrlFetchApp.fetch(url, options);
  //    var todayCovidtext = getCovidToday();
  //    if(flag==true){
  //      lineClient.replyMessage(token, { type: 'text', messages: todayCovidtext });
  //    }
    }
    
    catch (error) {
      Logger.log(error.name + "：" + error.message);
      return;
    }
      
    if (response.getResponseCode() === 200) {
      Logger.log("Sending message completed.");
    }
  }
  function sentNotifyDaily(){
    var ssId = "Your google sheet id.";
    var ss = SpreadsheetApp.openById(ssId);
    var sheetUsers = ss.getSheetByName("users");
    Logger = BetterLog.useSpreadsheet(ssId);
  
    var values = sheetUsers.getRange(1, 1, sheetUsers.getLastRow(),sheetUsers.getLastColumn()).getValues();
    var message = getCovidToday();
    var users = [];
    for(var i = 0;i<values.length; i++){
      users.push(values[i][0])
    }
    
    var header = {
      'Content-Type': 'application/json',
      'Authorization': "Bearer [Your token channel access token.]"
    };
    var body = {
      "to": users,
      "messages": [
      {
        "type":"text",
        "text":message
      },
      {
        "type":"text",
        "text": "************ วิธีใช้งาน ************\nสามารถเลือกแสดงตาม \"จังหวัด\" โดยพิมพ์ข้อมูล เช่น \"กรุงเทพมหานคร\" \nหรือต้องการเรียกดูทั้งหมด ให้พิมพ์ \"ล่าสุด\" หรือ \"ทั้งหมด\""
      }]
    }
    var options = {
      "method": "POST",
      "payload": JSON.stringify(body),
      "headers": header,
      };
    var url = "https://api.line.me/v2/bot/message/multicast";
    UrlFetchApp.fetch(url, options);
    Logger.log("Send Notify Daily Success.");
  }
  function getCovidSheet(){
    var ssId = "Your google sheet id.";
    var ss = SpreadsheetApp.openById(ssId);
    var sheetData = ss.getSheetByName("data1");
    var lastRow = sheetData.getLastRow();
    var msg = sheetData.getRange(lastRow, 1).getValue();
    return msg;
  }
  function getCovidToday() {
    var ssId = "Your google sheet id.";
    var ss = SpreadsheetApp.openById(ssId);
    var sheetData = ss.getSheetByName("data1");
    
    var url = 'https://covid19.th-stat.com/api/open/today'
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    var json = response.getContentText();
    var data = JSON.parse(json);
    var msg = "รายงานยอดติดเชื้อสะสม Covid-19 ในประเทศไทย update ทุกวันเวลา 12.00"
    +"\nUpdate ล่าสุดเมื่อ : "+data['UpdateDate']
    +"\nติดเชื้อสะสม : "+data['Confirmed']+" ราย"+" \nติดเชื้อเพิ่มขึ้น : "+ data['NewConfirmed']+" ราย"
    +"\nเสียชีวิตสะสม : "+data['Deaths']+" ราย"+" \nเสียชีวิตเพิ่มขึ้น : "+ data['NewDeaths']+" ราย"
    +"\nรักษาตัวในรพ.ตอนนี้ : "+data['Hospitalized']+" ราย"+" \nรักษาตัวในรพ.เพิ่มขึ้น : "+ data['NewHospitalized']+" ราย"
    +"\nหายแล้วตอนนี้ : "+data['Recovered']+" ราย"+" \nหายแล้วเพิ่มขึ้น : "+ data['NewHospitalized']+" ราย"
    ;
      
    var lastRow = sheetData.getLastRow();
    sheetData.getRange(lastRow + 1, 1).setValue(msg);
    return msg;
  }
  function getByProvince(reply){
    var ssId = "Your google sheet id.";
    var ss = SpreadsheetApp.openById(ssId);
    var sheetData = ss.getSheetByName("province");
    var num = 0;
    
    //Fix Bugs
    reply = reply.replace("จังหวัด","");
    reply = reply.trim();  
    
    if(reply === "กทม" || reply === "กรุงเทพ"){
      reply = "กรุงเทพมหานคร";
    }
      if(reply === "โคราช"){
      reply = "นครราชสีมา";
    }
    for(var i = 1;i<=77; i++){
      var province = sheetData.getRange(i, 1).getValue();
      if(reply.includes(province)){
         num = sheetData.getRange(i, 3).getValue();
        return province +" ติดเชื้อ : "+num+ " คน";
      }
      if(i==77){
        return "ขออภัย ไม่พบรายชื่อจังหวัด หรือ ป้อนข้อมูลผิด \nตัวอย่าง \"กรุงเทพมหานคร\""
      }
    }
    return num;
  }
  function setByProvince(){
    var ssId = "Your google sheet id.";
    var ss = SpreadsheetApp.openById(ssId);
    var sheetData = ss.getSheetByName("province");
    var data = getProvinceToday();
    for(var i = 1;i<=77; i++){
      var province = sheetData.getRange(i, 2).getValue();
      sheetData.getRange(i, 3).setValue(data[province]);
    }
  }
  function getProvinceToday(){
    var url = 'https://covid19.th-stat.com/api/open/cases/sum'
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    var json = response.getContentText();
    var data = JSON.parse(json);
    data = data['Province']
    return data
  }