function pluralizeday(num) {
  var forms = ['день', 'дня', 'дней'];
  var cases = [2, 0, 1, 1, 1, 2];
  var index = num % 100 > 4 && num % 100 < 20 ? 2 : cases[Math.min(num % 10, 5)];
  return num + ' ' + forms[index];
}

function getVerb(num) {
  var forms = ['остался', 'осталось', 'осталось'];
  var cases = [2, 0, 1, 1, 1, 2];
  var index = num % 100 > 4 && num % 100 < 20 ? 2 : cases[Math.min(num % 10, 5)];
  return forms[index];
}

function checkBalanceAndNotify() {
  var apiKey = apidasha;
  var apiUrl = "https://api.dashamail.ru/?method=account.get_balance&api_key=" + apiKey;

  // Отправка запроса к API
  var response = UrlFetchApp.fetch(apiUrl);
  var data = JSON.parse(response.getContentText());

  // Проверка типа ответа
  var responseType = data.response.msg.type;
  if (responseType === "message") {
    handleSuccessResponse(data.response.data);
  } else if (responseType === "error") {
    handleErrorResponse(data.response.msg);
  }
}

function handleSuccessResponse(data) {
  // Получение даты окончания тарифа
  var expirationDate = data.expiration_date;
  var parts = expirationDate.split(' ');
  var dateParts = parts[0].split('.');
  var timeParts = parts[1].split(':');
  var expirationDateTime = new Date(
    parseInt(dateParts[2]), // год
    parseInt(dateParts[1]) - 1, // месяц (начиная с 0)
    parseInt(dateParts[0]), // день
    parseInt(timeParts[0]), // часы
    parseInt(timeParts[1]), // минуты
    parseInt(timeParts[2]) // секунды
  );
  var daysRemaining = Math.ceil((expirationDateTime - new Date()) / (1000 * 60 * 60 * 24));
  var daysword = pluralizeday(daysRemaining)
  var daysword1 = getVerb(daysRemaining)

    var limitEmails = parseInt(data.limit_emails);
  var members = parseInt(data.members);

  // Проверка оставшегося количества дней и отправка сообщения в Telegram
  if (daysRemaining < 5 && limitEmails > members) {
    var telegramBotToken = API;
    var telegramChatId = defchatid;
    var message = "🤯😰 До даты окончания тарифа "+ daysword1 + " " + daysword +". Пожалуйста, оплатите тариф.";

    sendTelegramMessage(telegramBotToken, telegramChatId, message);
  }

  // Проверка лимита тарифа
  if (limitEmails < members) {
    var telegramBotToken = API;
    var telegramChatId = defchatid;
    var message = "☹️ Лимита тарифа не хватит, чтобы отправить сообщение на всю базу. Купите новый тариф.";

    sendTelegramMessage(telegramBotToken, telegramChatId, message);
  }
}

function handleErrorResponse(errorMsg) {
  if (errorMsg.err_code === 42) {
    var telegramBotToken = API;
    var telegramChatId = defchatid;
    var message = "❗Тариф закончился, письма не отправляются. Срочно пополните кабинет.";  
    sendTelegramMessage(telegramBotToken, telegramChatId, message);
    }
}

function sendTelegramMessage(botToken, chatId, message) {
var telegramApiUrl = "https://api.telegram.org/bot" + botToken + "/sendMessage";

var payload = {
method: "post",
contentType: "application/json",
payload: JSON.stringify({
chat_id: chatId,
text: message
})
};

// Отправка сообщения в Telegram
UrlFetchApp.fetch(telegramApiUrl, payload);
}
