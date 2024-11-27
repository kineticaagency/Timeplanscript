function TriggersAuto() {
  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры').getDataRange().getValues();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры');
  var startDate = sheet.getRange("D1").getValue();
  var endDate = sheet.getRange("F1").getValue();
  startDate = Utilities.formatDate(startDate, "Asia/Krasnoyarsk", "yyyy-MM-dd");
  endDate = Utilities.formatDate(endDate, "Asia/Krasnoyarsk", "yyyy-MM-dd");
  values.forEach( function(row, rowIndex) {
    if (rowIndex != 0 & rowIndex != 1) {
      var id = row[6];
      if (`${id}`.indexOf(',') > -1) {
        var ids = id. split(', ');
      } else { ids = [id]}
        var stroka = rowIndex + 1
        var query = "SELECT\nSUM(sent) as sent,\nSUM(clicked) as clicked,\nSUM(opened) as opened,\nminIf(sent_time, sent_time!='0000-00-00 00:00:00') as first_sent,\nMAX(open_time) as last_open,\nMAX(click_time) as last_click,\nSUM(unique_opened) as unique_opened,\nSUM(unique_clicked) as unique_clicked,\nSUM(unsubsribed) as unsubsribed,\nSUM(complained) as complained,\nSUM(preview) as preview,\nSUM (blk_spm) as spam_blocked,\nSUM (spm) as spam,\nSUM (blk) as blk,\nSUM (hrd) as hard,\nSUM (sft) as soft\nFROM\n(SELECT\nevent_time,\ngroupArray(event_time) as events_times,\narrayCount(status->status = 'SENT', groupArray(event_type)) as sent,\narrayCount(status->status = 'OPENED', groupArray(event_type)) as opened,\narrayFilter(time, status->(status = 'OPENED'), events_times, groupArray(event_type))[1] as open_time,\narrayFilter(time, status->(status = 'CLICKED'), events_times, groupArray(event_type))[1] as click_time,\narrayFilter(time, status->(status = 'SENT'), events_times, groupArray(event_type))[1] as sent_time,\narrayExists(status->(status = 'OPENED' OR status = 'CLICKED' OR status = 'UNSUBSCRIBED'), groupArray(event_type)) as unique_opened,\narrayCount(status->status = 'CLICKED', groupArray(event_type)) as clicked,\narrayExists(status->status = 'CLICKED', groupArray(event_type)) as unique_clicked,\narrayExists(status->status = 'UNSUBSCRIBED', groupArray(event_type)) as unsubsribed,\narrayExists(status->status = 'PREVIEW', groupArray(event_type)) as preview,\narrayExists(status->status = 'COMPLAINED', groupArray(event_type)) as complained,\narrayCount(status,cat,code->status = 'BOUNCED' AND cat = 'blk' AND code LIKE '3%', groupArray(event_type),   groupArray(bounce_category), groupArray(bounce_code)) as blk_spm,\narrayCount(status,cat,code->status = 'BOUNCED' AND cat = 'blk' AND code NOT LIKE '3%', groupArray(event_type),   groupArray(bounce_category), groupArray(bounce_code)) as blk,\narrayCount(status,cat,code->status = 'BOUNCED' AND cat = 'spm', groupArray(event_type),   groupArray(bounce_category), groupArray(bounce_code)) as spm,\narrayCount(status,cat,code->status = 'BOUNCED' AND cat = 'hrd', groupArray(event_type),   groupArray(bounce_category), groupArray(bounce_code)) as hrd,\narrayCount(status,cat,code->status = 'BOUNCED' AND cat = 'sft', groupArray(event_type),   groupArray(bounce_category), groupArray(bounce_code)) as sft\nFROM\ndm.raw_data\nWHERE\nevent_date >= '" + startDate +"'\nAND event_date <= '" + endDate + "'\nAND (";
        for (var i = 0; i < ids.length; i++) {
          query += "campaign_id = '" + ids[i] + "'";
          if (i < ids.length - 1) {
            query += " OR ";
          }
        }
        query += ") GROUP BY event_time)";

        Logger.log(query)

        var data = {
            "method": "raw.select",
            "query": query,
            "api_key": "5bc9306ba1e378ee0e31627245decd68"
           };

          var options = {
            'method': 'post',
            'contentType': 'application/json',
            'payload': JSON.stringify(data)
          };

          var answer = UrlFetchApp.fetch('https://api.dashamail.com/', options);

          var dann = JSON.parse(answer.getContentText());
          sheet.getRange(rowIndex+1, 14).setValue(dann.response.data[0].sent) // отправки
          sheet.getRange(rowIndex+1, 15).setValue(dann.response.data[0].unique_opened) // открытия
          sheet.getRange(rowIndex+1, 17).setValue(dann.response.data[0].unique_clicked) // клики
          sheet.getRange(rowIndex+1, 19).setValue(dann.response.data[0].unsubsribed) // отписки
          sheet.getRange(rowIndex+1, 21).setValue(dann.response.data[0].complained) // жалобы на спам
          sheet.getRange(rowIndex+1, 23).setValue(dann.response.data[0].preview) // превью
          sheet.getRange(rowIndex+1, 16).setValue('=O'+ stroka + '/N' + stroka)
          sheet.getRange(rowIndex+1, 18).setValue('=Q'+ stroka + '/N' + stroka)
          sheet.getRange(rowIndex+1, 20).setValue('=S'+ stroka + '/N' + stroka)
          sheet.getRange(rowIndex+1, 22).setValue('=U'+ stroka + '/N' + stroka)
      }
})
metrikatriggers()
}

function Test() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры');
  var values = sheet.getDataRange().getValues();
  values.forEach( function(row, rowIndex) {
    if (rowIndex != 0 & rowIndex != 1) {
      var utm = row[5];
      var prom = row[7]
      Logger.log(utm)
      if (prom != '') {
      Logger.log(prom) }
    }
})}

function metrikatriggers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры');
  var values = sheet.getDataRange().getValues();
  var startDate = sheet.getRange("D1").getValue();
  var endDate = sheet.getRange("F1").getValue();
  startDate = Utilities.formatDate(startDate, "Asia/Krasnoyarsk", "yyyy-MM-dd");
  endDate = Utilities.formatDate(endDate, "Asia/Krasnoyarsk", "yyyy-MM-dd");

  values.forEach( function(row, rowIndex) {
    if (rowIndex != 0 & rowIndex != 1) {
      var utm = row[5];
      var prom = row[7]

        // if (utm !== "" && utm.indexOf(",") == -1 && prom == "") {
        //   var headers = {
        //     "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
        //     "Authorization": "OAuth "
        //   }

        //   var options = {
        //     'method': 'get',
        //     'headers': headers,
        //     'redirect': 'follow'
        //   }

        //   var answer = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=6063082&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMCampaign=='"+utm+"'&date1="+ startDate +"&date2="+ endDate +"&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);

        //   stroka = rowIndex + 1;
        //   var dann = JSON.parse(answer.getContentText());
        //   income = Number(dann.totals[[0]]);
        //   purchase = Number(dann.totals[[1]]);
        //   sheet.getRange(rowIndex+1, 11).setValue(income) // деньги
        //   sheet.getRange(rowIndex+1, 10).setValue(purchase) // покупки
        // }

        if (utm !== "" && utm.indexOf(",") == -1 && prom !== "") {
          var headers = {
            "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
            "Authorization": "OAuth "
          }

          var options = {
            'method': 'get',
            'headers': headers,
            'redirect': 'follow',
            'muteHttpExceptions': true
          }

          var answer1 = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=6063082&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMSource!='email'&ym:s:productCoupon=='"+ prom +"'&date1="+ startDate +"&date2="+ endDate +"&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);

          Logger.log(prom);
          Logger.log(answer1.getContentText());

          var answer2 = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=6063082&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMCampaign=='"+utm+"'&date1="+ startDate +"&date2="+ endDate +"&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);

          stroka = rowIndex + 1;
          var dann = JSON.parse(answer1.getContentText());
          var dann2 = JSON.parse(answer2.getContentText());
          income = Number(dann.totals[[0]])+Number(dann2.totals[[0]]);
          purchase = Number(dann.totals[[1]]) + Number(dann2.totals[[1]]);
          sheet.getRange(stroka, 11).setValue(income) // деньги
          sheet.getRange(stroka, 10).setValue(purchase) // покупки
        }

        // if (utm !== "" && utm.indexOf(",") > -1 && prom !== "") {
        //   var ids = id. split(', ');
        //   for (var i = 0; i < ids.length; i++) {

        //   var headers = {
        //     "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
        //     "Authorization": "OAuth "
        //   }

        //   var options = {
        //     'method': 'get',
        //     'headers': headers,
        //     'redirect': 'follow',
        //     'muteHttpExceptions': true
        //   }

        //   var answer1 = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=6063082&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMSource!='email';ym:s:productCoupon=='"+ prom +"'&date1="+ startDate +"&date2="+ endDate +"&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);

        //   Logger.log(answer1.getContentText());

        //   var answer2 = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=6063082&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMCampaign=='"+utm+"'&date1="+ startDate +"&date2="+ endDate +"&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);

        //   stroka = rowIndex + 1;
        //   var dann = JSON.parse(answer1.getContentText());
        //   var dann2 = JSON.parse(answer2.getContentText());
        //   income = Number(dann.totals[[0]])+Number(dann2.totals[[0]]);
        //   purchase = Number(dann.totals[[1]]) + Number(dann2.totals[[1]]);
        //   sheet.getRange(stroka, 11).setValue(income) // деньги
        //   sheet.getRange(stroka, 10).setValue(purchase) // покупки
        // }}
      
    }
  })
}

function Test() {
  startdate("2022-10-01", "-1001502058030")
}

function startdate(text, chat_id) {
  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры').getDataRange().getValues();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры');
  var startDate = sheet.getRange("D1").setValue(text);
  send('И теперь дату окончания, пожалуйста', chat_id)
}

function Olegsay(text, chat_id) {
  var values = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры').getDataRange().getValues();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Триггеры');
  var startDate = sheet.getRange("F1").setValue(text);
  send('Спасибо, работаю над этим', chat_id)
  TriggersAuto()
  generateMessage(chat_id)
}

function generateMessage(chat_id) {
  // Get the sheet containing the data
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Триггеры");
  
  // Get the data range
  var dataRange = sheet.getDataRange();
  
  // Get the values in the range
  var data = dataRange.getValues();
  
  // Loop through each row of the data
  var message = "🙂 Итак, я собрал статистику, полюбуйтесь:\n\n";
  for (var i = 2; i < data.length; i++) {
    // Check if the row is empty
    if (data[i][0] == "") {
      break;
    }
    
    // Extract the mailing list name and distribution revenue
    var mailingListName = data[i][0];
    var distributionRevenue = data[i][10];
    var distributionPurshase = data[i][9];
    
    // Add the information to the message
    message += mailingListName + "\nДоход — " + distributionRevenue + "\nЗаказов — " + distributionPurshase + "\n\n";
  }
  
  send(message, chat_id)
}
