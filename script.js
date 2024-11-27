const defchatid = PropertiesService.getScriptProperties().getProperty("defchatid");
const idmetrika = PropertiesService.getScriptProperties().getProperty("idmetrika");
const apidasha = PropertiesService.getScriptProperties().getProperty("apidasha");
const link = PropertiesService.getScriptProperties().getProperty("link");
const mailService = PropertiesService.getScriptProperties().getProperty("mailService");
 
function doPost(e) {
  let tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Отправленные рассылки');
  let update = JSON.parse(e.postData.contents);
  var chat_id1 = update.message.chat.id;
  
  if (chat_id1 == defchatid) {
    if (update.hasOwnProperty('message')) {
      handleMessage(update.message, tab);
    }
  } else {
    let msg = update.message;
    let chat_id = msg.chat.id;
    let text = msg.text;
    if (text == 'Олег, отправил рассылку' && chat_id != defchatid) {
      send("Извините, я пока с такими чатами не работаю" + chat_id + defchatid, chat_id);
    }
  }
}

function handleMessage(msg, tab) {
  let chat_id = msg.chat.id;
  let text = msg.text;
  let user = msg.from.username;
  let monthRow = findMonthRow(tab);

  switch (text) {
    case "Олег, отправил рассылку":
    case "/send":
    case "/send@peredvizhnikleadplanbot":
      send("Круто, что вы смогли отправить рассылку! Сейчас я задам несколько вопросов по поводу этой рассылки. Отвечайте, пожалуйста, используя функцию Reply to.", chat_id);
      send("Назови, пожалуйста, тему рассылки.", chat_id);
      break;
    case "Олег, обнови показатели":
      send("Начинаю", chat_id);
      Olegobnovi(chat_id);
      break;
    case "Где деньги, Лебовски?":
      getMoney(chat_id);
      break;
    case "Олег, нужна статистика по триггерам":
      send("Без проблем. В ответ на это сообщение сообщи мне дату, от которой нужно начать расчет. В формате ГГГГ-ММ-ДД, пожалуйста", chat_id);
      break;
  }

  if (msg.hasOwnProperty('reply_to_message')) {
    handleReply(msg, tab, monthRow);
  }
}

function handleReply(msg, tab, monthRow) {
  let chat_id = msg.chat.id;
  let text = msg.text;
  let reply = msg.reply_to_message.text;
  let num = monthRow + 1;

  function getColumnByName(name) {
    let headers = tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0];
    return headers.indexOf(name) + 1;
  }

  switch (reply) {
    case "Назови, пожалуйста, тему рассылки.":
      insertRowAndSetValue(tab, num, getColumnByName("Тема письма"), text, chat_id, "Спасибо, записал! Какой был прехедер?");
      break;
    case "Без проблем. В ответ на это сообщение сообщи мне дату, от которой нужно начать расчет. В формате ГГГГ-ММ-ДД, пожалуйста":
      startdate(text, chat_id);
      break;
    case "И теперь дату окончания, пожалуйста":
      Olegsay(text, chat_id);
      break;
    case "Спасибо, записал! Какой был прехедер?":
      setRowValues(tab, num, getColumnByName("Прехедер"), text, chat_id, "Здорово! Если ты забыл сказать мне о рассылке заранее, и она была отправлена не сегодня, напиши дату, когда она была отправлена в формате ДД.ММ.ГГГГ. Если письмо было отправлено сегодня — ничего не пиши.");
      tab.getRange(num, getColumnByName("Дата")).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")).setFontWeight("normal");
      Utilities.sleep(3 * 1000)
      send("В ответ на этот вопрос пришли, пожалуйста, айди рассылки. Если их несколько — перечисли через запятую", chat_id)
      break;
    case "Здорово! Если ты забыл сказать мне о рассылке заранее, и она была отправлена не сегодня, напиши дату, когда она была отправлена в формате ДД.ММ.ГГГГ. Если письмо было отправлено сегодня — ничего не пиши.":
      setRowValues(tab, num, getColumnByName("Дата"), text, chat_id, "Окей, дату перезаписал");
      break;
    case "В ответ на этот вопрос пришли, пожалуйста, айди рассылки. Если их несколько — перечисли через запятую":
      setRowValues(tab, num, getColumnByName("Ссылка на письмо"), "https://lk.dashamail.ru/stat/preview.php?m=test&campaign=" + text.split(',')[0].trim(), chat_id, "Айди записал. Скажи номер задачи, в рамках которой работали над письмом 🫢");
      tab.getRange(num, getColumnByName("ID")).setValue(text).setFontWeight("normal");
      break;
    case "Айди записал. Скажи номер задачи, в рамках которой работали над письмом 🫢":
      send("Вы потратили на письмо 10 часов 🫨", chat_id);
      setRowValues(tab, num, getColumnByName("Номер задачи"), text, chat_id, "На какие сегменты отправляли рассылку?");
      break;
    case "На какие сегменты отправляли рассылку?":
      setRowValues(tab, num, getColumnByName("Сегмент"), text, chat_id, "Смело поделили! А кому не стали отправлять рассылку? Я, если что, про людей, которые выбрали частоту получения 😃");
      break;
    case "Смело поделили! А кому не стали отправлять рассылку? Я, если что, про людей, которые выбрали частоту получения 😃":
      setRowValues(tab, num, getColumnByName("Частота"), text, chat_id, "Ага, принял. Осталось два вопроса. Первый — от кого отправляли рассылку?");
      break;
    case "Ага, принял. Осталось два вопроса. Первый — от кого отправляли рассылку?":
      setRowValues(tab, num, getColumnByName("От кого"), text, chat_id, "Принял. И последний вопрос — какая UTM у рассылки? Если несколько — пиши через запятую.");
      break;
    case "Принял. И последний вопрос — какая UTM у рассылки? Если несколько — пиши через запятую.":
      setRowValues(tab, num, getColumnByName("UTM"), text, chat_id, "Отлично! Я все записал и еще несколько дней послежу за рассылкой 🙂");
      break;
    case "Олег, когда закончится тариф?":
      checkTariff(chat_id);
      break;
  }
}


function insertRowAndSetValue(sheet, row, col, value, chat_id, nextMessage) {
  sheet.insertRowAfter(row - 1);
  let cell = sheet.getRange(row, col);
  cell.setValue(value).setFontWeight("normal");
  send(nextMessage, chat_id);
}

function setRowValues(sheet, row, col, value, chat_id, nextMessage) {
  let cell = sheet.getRange(row, col);
  cell.setValue(value).setFontWeight("normal");
  send(nextMessage, chat_id);
}

function findMonthRow(sheet) {
  let lastRow = sheet.getLastRow();
  for (let i = 1; i <= lastRow; i++) {
    let cellValue = sheet.getRange(i, 1).getValue();
    if (typeof cellValue === 'string' && cellValue.includes('2024')) {
      return i;
    }
  }
  send(lastRow, "-1001500240126")
  return lastRow;
}



function getMoney(chat_id) {
  var headers = {
    "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
    "Authorization": "OAuth"
  };

  var options = {
    'method': 'get',
    'headers': headers,
    'redirect': 'follow'
  };

  var today = new Date();
  var firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  var date1 = Utilities.formatDate(firstDay, "GMT+0600", "yyyy-MM-dd");
  var answer = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=" + idmetrika + "&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMMedium=='email'&date1=" + date1 + "&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);
  var dann = JSON.parse(answer.getContentText());
  income = Number(dann.totals[[0]]);
  purchase = Number(dann.totals[[1]]);
  if (income != 0) {
    income = formatNumber(income);
    var purchaseText = pluralize(purchase);
    send("💸 Деньги — есть. В этом месяце пользователи совершили " + purchaseText + " на " + income + " ₽. Рад, что ваши письма приносят такие показатели 🥰", chat_id);
  } else {
    send("Денег нет. Но уверен, что виной всему ретроградный Меркурий, а не ваша работа ❤️", chat_id);
  }
}

function checkTariff(chat_id) {
  var apiUrl = "https://api.dashamail.ru/?method=account.get_balance&api_key=" + apidasha;

  var response = UrlFetchApp.fetch(apiUrl);
  var data = JSON.parse(response.getContentText());
  var expirationDate = data.response.data.expiration_date;
  var parts = expirationDate.split(' ');
  var dateParts = parts[0].split('.');
  var timeParts = parts[1].split(':');
  var expirationDateTime = new Date(
    parseInt(dateParts[2]),
    parseInt(dateParts[1]) - 1,
    parseInt(dateParts[0]),
    parseInt(timeParts[0]),
    parseInt(timeParts[1]),
    parseInt(timeParts[2])
  );
  var daysRemaining = Math.ceil((expirationDateTime - new Date()) / (1000 * 60 * 60 * 24));
  var daysword = pluralizeday(daysRemaining);
  var daysword1 = getVerb(daysRemaining);
  var message = "До даты окончания тарифа " + daysword1 + " " + daysword;
  send(message, chat_id);
}

function getMoney(chat_id) {
  var headers = {
    "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
    "Authorization": "OAuth"
  };

  var options = {
    'method': 'get',
    'headers': headers,
    'redirect': 'follow'
  };

  var today = new Date();
  var firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  var date1 = Utilities.formatDate(firstDay, "GMT+0600", "yyyy-MM-dd");
  var answer = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=" + idmetrika + "&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMMedium=='email'&date1=" + date1 + "&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);
  var dann = JSON.parse(answer.getContentText());
  income = Number(dann.totals[[0]]);
  purchase = Number(dann.totals[[1]]);
  if (income != 0) {
    income = formatNumber(income);
    var purchaseText = pluralize(purchase);
    send("💸 Деньги — есть. В этом месяце пользователи совершили " + purchaseText + " на " + income + " ₽. Рад, что ваши письма приносят такие показатели 🥰", chat_id);
  } else {
    send("Денег нет. Но уверен, что виной всему ретроградный Меркурий, а не ваша работа ❤️", chat_id);
  }
}

function checkTariff(chat_id) {
  var apiUrl = "https://api.dashamail.ru/?method=account.get_balance&api_key=" + apidasha;

  var response = UrlFetchApp.fetch(apiUrl);
  var data = JSON.parse(response.getContentText());
  var expirationDate = data.response.data.expiration_date;
  var parts = expirationDate.split(' ');
  var dateParts = parts[0].split('.');
  var timeParts = parts[1].split(':');
  var expirationDateTime = new Date(
    parseInt(dateParts[2]),
    parseInt(dateParts[1]) - 1,
    parseInt(dateParts[0]),
    parseInt(timeParts[0]),
    parseInt(timeParts[1]),
    parseInt(timeParts[2])
  );
  var daysRemaining = Math.ceil((expirationDateTime - new Date()) / (1000 * 60 * 60 * 24));
  var daysword = pluralizeday(daysRemaining);
  var daysword1 = getVerb(daysRemaining);
  var message = "До даты окончания тарифа " + daysword1 + " " + daysword;
  send(message, chat_id);
}

function getMoney(chat_id) {
  var headers = {
    "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
    "Authorization": "OAuth"
  };

  var options = {
    'method': 'get',
    'headers': headers,
    'redirect': 'follow'
  };

  var today = new Date();
  var firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  var date1 = Utilities.formatDate(firstDay, "GMT+0600", "yyyy-MM-dd");
  var answer = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=" + idmetrika + "&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMMedium=='email'&date1=" + date1 + "&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);
  var dann = JSON.parse(answer.getContentText());
  income = Number(dann.totals[[0]]);
  purchase = Number(dann.totals[[1]]);
  if (income != 0) {
    income = formatNumber(income);
    var purchaseText = pluralize(purchase);
    send("💸 Деньги — есть. В этом месяце пользователи совершили " + purchaseText + " на " + income + " ₽. Рад, что ваши письма приносят такие показатели 🥰", chat_id);
  } else {
    send("Денег нет. Но уверен, что виной всему ретроградный Меркурий, а не ваша работа ❤️", chat_id);
  }
}

function checkTariff(chat_id) {
  var apiUrl = "https://api.dashamail.ru/?method=account.get_balance&api_key=" + apidasha;

  var response = UrlFetchApp.fetch(apiUrl);
  var data = JSON.parse(response.getContentText());
  var expirationDate = data.response.data.expiration_date;
  var parts = expirationDate.split(' ');
  var dateParts = parts[0].split('.');
  var timeParts = parts[1].split(':');
  var expirationDateTime = new Date(
    parseInt(dateParts[2]),
    parseInt(dateParts[1]) - 1,
    parseInt(dateParts[0]),
    parseInt(timeParts[0]),
    parseInt(timeParts[1]),
    parseInt(timeParts[2])
  );
  var daysRemaining = Math.ceil((expirationDateTime - new Date()) / (1000 * 60 * 60 * 24));
  var daysword = pluralizeday(daysRemaining);
  var daysword1 = getVerb(daysRemaining);
  var message = "До даты окончания тарифа " + daysword1 + " " + daysword;
  send(message, chat_id);
}


function getMoney(chat_id) {
  var headers = {
    "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
    "Authorization": "OAuth"
  };

  var options = {
    'method': 'get',
    'headers': headers,
    'redirect': 'follow'
  };

  var today = new Date();
  var firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  var date1 = Utilities.formatDate(firstDay, "GMT+0600", "yyyy-MM-dd");
  var answer = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id=" + idmetrika + "&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMMedium=='email'&date1=" + date1 + "&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);
  var dann = JSON.parse(answer.getContentText());
  income = Number(dann.totals[[0]]);
  purchase = Number(dann.totals[[1]]);
  if (income != 0) {
    income = formatNumber(income);
    var purchaseText = pluralize(purchase);
    send("💸 Деньги — есть. В этом месяце пользователи совершили " + purchaseText + " на " + income + " ₽. Рад, что ваши письма приносят такие показатели 🥰", chat_id);
  } else {
    send("Денег нет. Но уверен, что виной всему ретроградный Меркурий, а не ваша работа ❤️", chat_id);
  }
}

function checkTariff(chat_id) {
  var apiUrl = "https://api.dashamail.ru/?method=account.get_balance&api_key=" + apidasha;

  var response = UrlFetchApp.fetch(apiUrl);
  var data = JSON.parse(response.getContentText());
  var expirationDate = data.response.data.expiration_date;
  var parts = expirationDate.split(' ');
  var dateParts = parts[0].split('.');
  var timeParts = parts[1].split(':');
  var expirationDateTime = new Date(
    parseInt(dateParts[2]),
    parseInt(dateParts[1]) - 1,
    parseInt(dateParts[0]),
    parseInt(timeParts[0]),
    parseInt(timeParts[1]),
    parseInt(timeParts[2])
  );
  var daysRemaining = Math.ceil((expirationDateTime - new Date()) / (1000 * 60 * 60 * 24));
  var daysword = pluralizeday(daysRemaining);
  var daysword1 = getVerb(daysRemaining);
  var message = "До даты окончания тарифа " + daysword1 + " " + daysword;
  send(message, chat_id);
}

function send(msg, chat_id) {
  var payload = {
    'chat_id': chat_id,
    'message': msg
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch('https://amp.kinetica.su/timeplantable/olegsend.php', options);
}



function autodate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Отправленные рассылки');
  var values = sheet.getDataRange().getValues();
  
  var headers = values[0];
  
  var columnIndexes = {
    id: headers.indexOf("ID"),
    datecheck: headers.indexOf("Дата"),
    sent: headers.indexOf("Отправок"),
    opened: headers.indexOf("Открытий"),
    openRate: headers.indexOf("Open Rate"),
    clicked: headers.indexOf("Кликов"),
    clickRate: headers.indexOf("Click Rate"),
    unsubscribed: headers.indexOf("Отписок"),
    unsubRate: headers.indexOf("Unsub Rate"),
    complained: headers.indexOf("Абьюзов"),
    abuseRate: headers.indexOf("Абьюз Rate"),
    preview: headers.indexOf("Превью"),
    bounced: headers.indexOf("Баунсов"),
    bouncedRate: headers.indexOf("Баунс Rate"),
  };

  var fiveMonthsAgo = new Date();
  fiveMonthsAgo.setMonth(fiveMonthsAgo.getMonth() - 3);

  values.forEach(function(row, rowIndex) {
    if (rowIndex > 1) {
      var id = row[columnIndexes.id];
      var datecheck = row[columnIndexes.datecheck];

      if (datecheck > fiveMonthsAgo.getTime()) {
        var ids = `${id}`.indexOf(',') > -1 ? id.split(', ') : [id];
        var stats = { sent: 0, opened: 0, clicked: 0, unsubscribed: 0, complained: 0, preview: 0, bounced: 0 };

        ids.forEach(function(item) {
          var options = { 'method': 'get', 'redirect': 'follow' };
          var answer = UrlFetchApp.fetch('https://api.dashamail.com?api_key=' + apidasha + '&campaign_id=' + item + '&method=reports.summary', options);
          var data = JSON.parse(answer.getContentText()).response.data;
          stats.sent += Number(data.sent);
          stats.opened += Number(data.unique_opened);
          stats.clicked += Number(data.unique_clicked);
          stats.unsubscribed += Number(data.unsubscribed);
          stats.complained += Number(data.complained);
          stats.preview += Number(data.preview);
          
          // Суммирование blk + hard + soft для bounced
          stats.bounced += Number(data.blk) + Number(data.hard) + Number(data.soft);
        });

        // Обновление значений в ячейках
        Object.keys(stats).forEach(function(stat) {
          var index = columnIndexes[stat];
          if (index !== -1) {
            sheet.getRange(rowIndex + 1, index + 1).setValue(stats[stat]);
          }
        });

        // Формулы для процентов
        var setFormula = function(numeratorIndex, denominatorIndex, formulaIndex) {
          if (numeratorIndex !== -1 && denominatorIndex !== -1 && formulaIndex !== -1) {
            sheet.getRange(rowIndex + 1, formulaIndex + 1).setFormula(`=${sheet.getRange(rowIndex + 1, numeratorIndex + 1).getA1Notation()}/${sheet.getRange(rowIndex + 1, denominatorIndex + 1).getA1Notation()}`);
          }
        };
        
        setFormula(columnIndexes.opened, columnIndexes.sent, columnIndexes.openRate);
        setFormula(columnIndexes.clicked, columnIndexes.sent, columnIndexes.clickRate);
        setFormula(columnIndexes.unsubscribed, columnIndexes.sent, columnIndexes.unsubRate);
        setFormula(columnIndexes.complained, columnIndexes.sent, columnIndexes.abuseRate);
        setFormula(columnIndexes.bounced, columnIndexes.sent, columnIndexes.bouncedRate);
      }
    }
  });

  // Форматирование ячеек
  var formatColumn = function(index, format) {
    if (index !== -1) {
      sheet.getRange(2, index + 1, sheet.getLastRow() - 1).setNumberFormat(format);
    }
  };

  // Форматирование для процентов
  formatColumn(columnIndexes.openRate, '0.00%');
  formatColumn(columnIndexes.clickRate, '0.00%');
  formatColumn(columnIndexes.unsubRate, '0.00%');
  formatColumn(columnIndexes.abuseRate, '0.00%');
  formatColumn(columnIndexes.bouncedRate, '0.00%');

  // Форматирование для чисел
  formatColumn(columnIndexes.sent, '[<10000]0;[>=10000]#,###');
  formatColumn(columnIndexes.opened, '[<10000]0;[>=10000]#,###');
  formatColumn(columnIndexes.clicked, '[<10000]0;[>=10000]#,###');
  formatColumn(columnIndexes.unsubscribed, '[<10000]0;[>=10000]#,###');
  formatColumn(columnIndexes.complained, '[<10000]0;[>=10000]#,###');
  formatColumn(columnIndexes.preview, '[<10000]0;[>=10000]#,###');
  formatColumn(columnIndexes.bounced, '[<10000]0;[>=10000]#,###');
}



function metrika() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Отправленные рассылки');
  var values = sheet.getDataRange().getValues();

  // Форматирование ячеек
  var formatColumn = function(index, format) {
    if (index !== -1) {
      sheet.getRange(2, index + 1, sheet.getLastRow() - 1).setNumberFormat(format);
    }
  };
  
  // Заголовки столбцов
  var headers = values[0]; // Строка с заголовками - она всегда на первом месте

  // Определение индексов нужных столбцов по названию
  var columnIndexes = {
    utm: headers.indexOf("UTM"),
    datecheck: headers.indexOf("Дата"),
    orders: headers.indexOf("Заказов"),  // Количество заказов
    income: headers.indexOf("Доход"),   // Доход
    clicks: headers.indexOf("Кликов"),   // Количество кликов
    cost: headers.indexOf("Стоимость"),  // Стоимость
    cpa: headers.indexOf("CPA"),
    cr: headers.indexOf("CR")
  };

  // Найдем самую раннюю дату в таблице
  var earliestDate = new Date();
  values.slice(1).forEach(function(row) {
    var rowDate = new Date(row[columnIndexes.datecheck]);
    if (rowDate < earliestDate) {
      earliestDate = rowDate;
    }
  });
  
  // Преобразуем дату в формат YYYY-MM-DD для использования в запросах
  var formattedDate = Utilities.formatDate(earliestDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  values.forEach(function(row, rowIndex) {
    if (rowIndex !== 0) { // Пропускаем заголовки
      var utm = row[columnIndexes.utm];
      var datecheck = row[columnIndexes.datecheck];

      var fiveMonthsAgo = new Date();
      fiveMonthsAgo.setMonth(fiveMonthsAgo.getMonth() - 3); // Для фильтрации по дате

      if (utm && datecheck > fiveMonthsAgo) {  // Убедимся, что дата после самой ранней
        var utmValues = utm.indexOf(',') > -1 ? utm.split(', ') : [utm]; // Поддержка нескольких UTM

        utmValues.forEach(function(item) {
          var headers = {
            "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
            "Authorization": "OAuth"
          };

          var options = {
            'method': 'get',
            'headers': headers,
            'redirect': 'follow'
          };

          var apiUrl = "https://api-metrika.yandex.net/stat/v1/data/bytime?id=" + idmetrika +
            "&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases" +
            "&filters=ym:s:cross_device_last_significantUTMCampaign=='" + item + "'" +
            "&date1=" + formattedDate + // Используем самую раннюю дату
            "&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true";

          var answer = UrlFetchApp.fetch(apiUrl, options);
          var dann = JSON.parse(answer.getContentText());
          
          var totalIncome = Number(dann.totals[0]);
          var totalPurchase = Number(dann.totals[1]);
          
          var rowIndex1 = rowIndex + 1;
          
          // Обновление значений для столбцов "Заказы" и "Доход"
          sheet.getRange(rowIndex1, columnIndexes.orders + 1).setValue(totalPurchase);
          sheet.getRange(rowIndex1, columnIndexes.income + 1).setValue(totalIncome);
          
          // Установим формулы для "CR" и "CPA" в правильные столбцы
          if (columnIndexes.clicks !== -1) {
            // Для CR формула: Заказы / Кликов
            var crFormula = '=IF(' + sheet.getRange(rowIndex1, columnIndexes.clicks + 1).getA1Notation() + '>0; ' + sheet.getRange(rowIndex1, columnIndexes.orders + 1).getA1Notation() + ' / ' + sheet.getRange(rowIndex1, columnIndexes.clicks + 1).getA1Notation() + '; 0)';
            sheet.getRange(rowIndex1, columnIndexes.cr + 1 ).setFormula(crFormula); // CR в соседний столбец от Заказов
          }

          if (columnIndexes.cost !== -1) {
            // Для CPA формула: Доход / Стоимость
            var cpaFormula = '=IF(' + sheet.getRange(rowIndex1, columnIndexes.cost + 1).getA1Notation() + '>0; ' + sheet.getRange(rowIndex1, columnIndexes.income + 1).getA1Notation() + ' / ' + sheet.getRange(rowIndex1, columnIndexes.cost + 1).getA1Notation() + '; 0)';
            sheet.getRange(rowIndex1, columnIndexes.cpa + 1).setFormula(cpaFormula); // CPA в соседний столбец от Стоимости
          }

          formatColumn(columnIndexes.orders, '[<10000]0;[>=10000]#,###');
          formatColumn(columnIndexes.income, '[<10000]0 ₽;[>=10000]#,### ₽');
          formatColumn(columnIndexes.cpa, '[<10000]0 ₽;[>=10000]#,### ₽');
          formatColumn(columnIndexes.cr, '0.00%');
        });
      }
    }
  });
}






function Olegsend() {
  // autodate();
  // metrika();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Отправленные рассылки');
  var values = sheet.getDataRange().getValues();
  
  // Получаем заголовки из первой строки
  var headers = values[0];
  
  // Создаем объект для хранения индексов столбцов по названиям
  var columnIndexes = {
    datepismo: headers.indexOf("Дата"),
    subject: headers.indexOf("Тема письма"),
    link: headers.indexOf("Ссылка на письмо"),
    income: headers.indexOf("Доход"),
    sended: headers.indexOf("Отправок"),
    open: headers.indexOf("Открытий"),
    openrate: headers.indexOf("Open Rate"),
    click: headers.indexOf("Кликов"),
    clickrate: headers.indexOf("Click Rate"),
    unsub: headers.indexOf("Отписок"),
    unsubrate: headers.indexOf("Unsub Rate"),
    spam: headers.indexOf("Абьюзов"),
    spamrate: headers.indexOf("Абьюз Rate"),
    orderrate: headers.indexOf("CR"),
    order: headers.indexOf("Заказов"),
    campaignId: headers.indexOf("ID"),
    bounce: headers.indexOf("Баунсов"),
    bouncerate: headers.indexOf("Баунс Rate")
  };
  
  values.forEach(function(row, rowIndex) {
    if (rowIndex !== 0) { // Пропускаем заголовки
      var datepismo = row[columnIndexes.datepismo];
      let d = new Date();
      var data = d.setDate(d.getDate() - 5); // Дата 5 дней назад
      if (datepismo >= data) {
        var income = columnIndexes.income !== -1 ? formatNumber(row[columnIndexes.income]) : '';
        var subject = columnIndexes.subject !== -1 ? row[columnIndexes.subject] : '';
        var sended = columnIndexes.sended !== -1 ? formatNumber(row[columnIndexes.sended]) : '';
        var open = columnIndexes.open !== -1 ? formatNumber(row[columnIndexes.open]) : '';
        var openrate = columnIndexes.openrate !== -1 ? (row[columnIndexes.openrate] * 100).toFixed(2) : '';
        var click = columnIndexes.click !== -1 ? formatNumber(row[columnIndexes.click]) : '';
        var clickrate = columnIndexes.clickrate !== -1 ? (row[columnIndexes.clickrate] * 100).toFixed(2) : '';
        var unsub = columnIndexes.unsub !== -1 ? formatNumber(row[columnIndexes.unsub]) : '';
        var unsubrate = columnIndexes.unsubrate !== -1 ? (row[columnIndexes.unsubrate] * 100).toFixed(2) : '';
        var spam = columnIndexes.spam !== -1 ? formatNumber(row[columnIndexes.spam]) : '';
        var spamrate = columnIndexes.spamrate !== -1 ? (row[columnIndexes.spamrate] * 100).toFixed(2) : '';
        var orderrate = columnIndexes.orderrate !== -1 ? (row[columnIndexes.orderrate] * 100).toFixed(2) : '';
        var order = columnIndexes.order !== -1 ? formatNumber(row[columnIndexes.order]) : '';
        var bounce = columnIndexes.bounce !== -1 ? formatNumber(row[columnIndexes.bounce]) : '';
        var bouncerate = columnIndexes.bouncerate !== -1 ? (row[columnIndexes.bouncerate] * 100).toFixed(2) : '';

        // Запрос к API dashamail.com
        var campaignId = row[columnIndexes.campaignId];
        var apiUrl = "https://api.dashamail.com/?api_key=" + apidasha + "&method=raw.select&query=SELECT url_original, COUNT(DISTINCT email) as quantity FROM dm.raw_data WHERE event_type = 'CLICKED' AND campaign_id IN (" + campaignId + ") GROUP BY url_original ORDER BY quantity DESC LIMIT 3";
        var response = UrlFetchApp.fetch(apiUrl);
        var responseData = JSON.parse(response.getContentText());

        // Формирование сообщения с самыми прокликиваемыми ссылками
        var topLinksMsg = "А вот самые прокликиваемые ссылки:\n";
        responseData.response.data.forEach(function(linkData) {
          topLinksMsg += removeUTMParams(linkData.url_original) + " — " + linkData.quantity + "\n";
        });

        // Формирование и отправка сообщения
        var message = "Йоп, всем привет! Помните, отправляли рассылку с темой <a href=\"" + row[columnIndexes.link] + "\">«" + subject + "»</a>? Принес ее результаты\n\n";
        
        if (sended) message += "Отправок — " + sended + "\n";
        if (open) message += "Открытий — " + open + " (" + openrate + "%)\n";
        if (click) message += "Кликов — " + click + " (" + clickrate + "%)\n";
        if (unsub) message += "Отписок — " + unsub + " (" + unsubrate + "%)\n";
        if (spam) message += "Жалоб на спам — " + spam + " (" + spamrate + "%)\n";
        if (bounce) message += "Возвратов — " + bounce + " (" + bouncerate + "%)\n";
        if (order) message += "\nЗаказов — " + order + " (" + orderrate + "%)\n";
        if (income) message += "Доход — " + income + " ₽\n";
        
        message += "\n" + topLinksMsg + "\nНадеюсь, показатели вам понравились. Хорошего вечера! ❤";
        send(message, defchatid);
      }
    }
  });
}



function removeUTMParams(url) {
  // Используем регулярные выражения для удаления параметров из URL
  return url.replace(/(\?|&)utm_campaign=[^&]+/g, '')
            .replace(/(\?|&)utm_medium=[^&]+/g, '')
            .replace(/(\?|&)utm_source=[^&]+/g, '')
            .replace(/(\?|&)utm_term=[^&]+/g, '')
            .replace(/(\?|&)roistat=[^&]+/g, '')
            .replace(/&$/, '') // Удаляем лишние амперсанды в конце URL, если они есть
            .replace(/(\?|&)$/, ''); // Удаляем вопросительный знак или амперсанд в конце URL, если они есть
}

function pluralize(num) {
  var forms = ['покупка', 'покупки', 'покупок'];
  var cases = [2, 0, 1, 1, 1, 2];
  var index = num % 100 > 4 && num % 100 < 20 ? 2 : cases[Math.min(num % 10, 5)];
  return num + ' ' + forms[index];
}

function Olegobnovi(chat_id) {
  try {
        send("Я в начале функции", chat_id)
        autodate();
        send("Я перехожу к метрике", chat_id)
        metrika();
        send("Я закончил", chat_id)
        send("Уже сделал. Проверяй в таблице: "+ link +" ❤", chat_id)
        } catch { send("Произошла чудовищная ошибка и часть данных не изменилась. Скорее всего, некоторые рассылки были внесены в тайм-план, но еще не отправились 😥", chat_id)}
}

function formatNumber(number) {
  if (number < 10000) {
    return number.toString();
  } else {
    return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  }
}
