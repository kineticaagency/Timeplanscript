function setSettings(settings) {
    PropertiesService.getScriptProperties().setProperties({
        "API": settings.API,
        "defchatid": settings.defchatid,
        "idmetrika": settings.idmetrika,
        "apidasha": settings.apidasha,
        "link": settings.link,
        "mailService": settings.mailService,
        "analyticsService": settings.analyticsService
    }, true);
}

function addToUi() {
    var user = Session.getActiveUser().getEmail();
    var domain = user.split('@')[1];
    var owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();

    if (domain === 'kinetica.su' || user === 'kinetica.wm' || user === owner) {
        var html = HtmlService.createHtmlOutputFromFile('Settings')
            .setWidth(400)
            .setHeight(500);
        SpreadsheetApp.getUi().showModalDialog(html, 'Настройки');
    } else {
        SpreadsheetApp.getUi().alert('У вас нет доступа к настройкам.');
    }
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Пользовательские настройки')
        .addItem('Настройки', 'addToUi')
        .addToUi();
}

function checkAndSetupEnvironment() {
    // Установка требуемых триггеров
    setupTriggers();
}

function setupTriggers() {
    // Необходимые триггеры и их параметры
    var requiredTriggers = [
        {functionName: 'moneyWhere', time: {hour: 9}, type: 'time-driven', frequency: 'daily'},
        {functionName: 'updateMonthlyData', time: {hour: 2, date: 1}, type: 'time-driven', frequency: 'monthly'},
        {functionName: 'checkBalanceAndNotify', time: {hour: 9}, type: 'time-driven', frequency: 'daily'},
        {functionName: 'autodate', type: 'time-driven', frequency: 'every12hours'},
        {functionName: 'metrika', type: 'time-driven', frequency: 'every12hours'},
        {functionName: 'OlegSend', time: {hour: 16}, type: 'time-driven', frequency: 'daily'}
    ];

    // Удаляем все существующие триггеры и создаем их заново, чтобы избежать дублирования
    deleteAllTriggers();
    requiredTriggers.forEach(createTrigger);
}

function deleteAllTriggers() {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
        ScriptApp.deleteTrigger(allTriggers[i]);
    }
}

function createTrigger(trigger) {
    var builder = ScriptApp.newTrigger(trigger.functionName);
    switch (trigger.type) {
        case 'time-driven':
            switch (trigger.frequency) {
                case 'daily':
                    if (trigger.time.hour !== undefined) builder.timeBased().atHour(trigger.time.hour).everyDays(1).inTimezone(Session.getScriptTimeZone()).create();
                    break;
                case 'monthly':
                    if (trigger.time.hour !== undefined && trigger.time.date !== undefined) builder.timeBased().atHour(trigger.time.hour).onMonthDay(trigger.time.date).inTimezone(Session.getScriptTimeZone()).create();
                    break;
                case 'every12hours':
                    builder.timeBased().everyHours(12).create();
                    break;
            }
            break;
    }
}

function getCurrentTimeStatus() {
    var serverTime = new Date();
    var gmt7Time = new Date(serverTime.toLocaleString("en-US", {timeZone: "Asia/Novosibirsk"}));

    var timeDifference = Math.abs(serverTime - gmt7Time);
    var isTimeZoneCorrect = timeDifference < 60 * 1000; // Разница менее одной минуты

    return isTimeZoneCorrect ? "GMT+7" : "incorrect";
}

function getSettings() {
    return PropertiesService.getScriptProperties().getProperties();
}
