function updateMonthlyData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var date = new Date();
  
  // Получаем название текущего месяца и год
  var month = getMonthName(date.getMonth()) + " " + date.getFullYear();
  
  // Определяем строку, в которую вставить новый месяц
  var insertRow = findInsertRow(sheet, month);
  
  // Вставляем название месяца с годом в нужную строку
  sheet.insertRowsBefore(insertRow, 1);  // Вставляем новую строку перед предыдущим месяцем
  sheet.getRange(insertRow, 1).setValue(month).setFontWeight("bold").setFontSize(12);
  
  // Удаляем месяцы без рассылок, кроме текущего
  removeEmptyMonths(sheet, date);
}

// Возвращает название месяца
function getMonthName(month) {
  var monthNames = [    
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", 
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"  
  ];
  return monthNames[month];
}

// Находит строку, в которую нужно вставить новый месяц
function findInsertRow(sheet, newMonth) {
  let lastRow = sheet.getLastRow();
  let monthNames = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", 
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
  ];
  
  for (let i = 1; i <= lastRow; i++) {
    let cellValue = sheet.getRange(i, 1).getValue();
    
    if (typeof cellValue === 'string') {
      let cellMonth = cellValue.split(' ')[0];  // Получаем название месяца из строки
      let cellYear = cellValue.split(' ')[1];   // Получаем год из строки
      
      // Если текущий месяц предшествует вставляемому месяцу, возвращаем эту строку
      if (monthNames.indexOf(cellMonth) < monthNames.indexOf(newMonth.split(' ')[0]) 
          && cellYear === newMonth.split(' ')[1]) {
        return i;
      }
    }
  }
  
  // Если не нашли строку, возвращаем последнюю строку + 1
  return lastRow + 1;
}

// Функция для удаления месяцев без рассылок
function removeEmptyMonths(sheet, currentDate) {
  var lastRow = sheet.getLastRow();
  var currentMonth = getMonthName(currentDate.getMonth()) + " " + currentDate.getFullYear();
  
  // Проходимся по всем строкам с конца, чтобы корректно удалять
  for (let i = lastRow; i >= 1; i--) {
    let cellValue = sheet.getRange(i, 1).getValue();
    
    // Проверяем только строки с названиями месяцев
    if (typeof cellValue === 'string' && cellValue.match(/^[А-Яа-я]+\s\d{4}$/)) {
      let monthStartRow = i;
      let nextMonthRow = findNextMonthRow(sheet, monthStartRow, lastRow);
      
      // Проверяем, есть ли рассылки между текущим месяцем и следующим месяцем
      let hasData = false;
      for (let j = monthStartRow + 1; j < nextMonthRow; j++) {
        let rowValues = sheet.getRange(j, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (rowValues.some(value => value !== "")) {
          hasData = true;
          break;
        }
      }
      
      // Удаляем месяц, если данных нет и это не текущий месяц
      if (!hasData && cellValue !== currentMonth) {
        sheet.deleteRow(monthStartRow);
      }
    }
  }
}

// Находит строку следующего месяца, начиная с текущей строки
function findNextMonthRow(sheet, startRow, lastRow) {
  for (let i = startRow + 1; i <= lastRow; i++) {
    let cellValue = sheet.getRange(i, 1).getValue();
    if (typeof cellValue === 'string' && cellValue.match(/^[А-Яа-я]+\s\d{4}$/)) {
      return i;
    }
  }
  return lastRow + 1;  // Если следующий месяц не найден, возвращаем последнюю строку
}
