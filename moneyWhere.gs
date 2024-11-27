function moneyWhere() {
      var chat_id = defchatid;
      var headers = {
              "Cookie": "JSESSIONID=node0ijangw7bbymi1kfdubsb7gxkl13117369.node0",
              "Authorization": "OAuth "
            }

            var options = {
              'method': 'get',
              'headers': headers,
              'redirect': 'follow'
            }

            var today = new Date(); // Get the current date
            var firstDay = new Date(today.getFullYear(), today.getMonth(), 1); // Get the first day of the current month
        var date1 = Utilities.formatDate(firstDay, "GMT+0600", "yyyy-MM-dd"); // Format first day as 2022-01-01
        var answer = UrlFetchApp.fetch("https://api-metrika.yandex.net/stat/v1/data/bytime?id="+idmetrika+"&metrics=ym:s:ecommerceRevenue,ym:s:ecommercePurchases&filters=ym:s:cross_device_last_significantUTMMedium=='email'&date1=" + date1+"&group=all&accuracy=1&attribution=cross_device_last_significant&cross_device=true", options);
        var dann = JSON.parse(answer.getContentText());
            income = Number(dann.totals[[0]]);
            purchase = Number(dann.totals[[1]]);
            if (income != 0) {
            income = formatNumber(income)
        var purchaseText = pluralize(purchase);
        send("💸 Деньги — есть. В этом месяце пользователи совершили " + purchaseText + " на " + income + " ₽. Рад, что ваши письма приносят такие показатели 🥰", chat_id)} else {
          send("Денег нет. Но уверен, что виной всему ретроградный Меркурий, а не ваша работа ❤️", chat_id)
        }
}
