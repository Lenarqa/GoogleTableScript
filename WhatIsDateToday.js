function myFunction() {
  
  var parityWeek;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();//сегодняшний день месяц год
  var day = now.getDate();//сегодняшний день только дата без месяца и года  
  var numDayWeek = now.getDay(); //день недели от 0(воскресенье) до 6(суббота);
  var weekDayArr = ["Вс","Пн","Вт","Ср","Чт","Пт","Сб"];//Обязательно нужно начинать этот массив с Вс т.к getDay() начинается с воскресенья.
  
  var dayWeek1 = ["B","C","D","E","F","G","H"];//массив букв столбиков 1 недели
  var dayWeek2 = ["I","J","K","L","M","N","O"];//массив букв столбиков 2 недели
  

  //Проверки значений
  //ss.getActiveSheet().getRange("B1").setValue(now);
  //ss.getActiveSheet().getRange("C1").setValue(day);
  //ss.getActiveSheet().getRange("D1").setValue(numDayWeek);
  //ss.getActiveSheet().getRange("E1").setValue(weekDayArr[numDayWeek]);
 
  //получаем значения дней недели в таблице в виде стринга для сравнения.
  var date1 = ss.getRange("B4:H4").getValues();
  // Для этого указываем диапазон клеток дней недели четной недели
  
  //сдесь проверяем четная или нечетная сейчас неделя.
  /*
  Пн  Вт  Ср  Чт  Пт  Сб  Вс
  1   2   3   4   5   6   0   numDayWeek
  18 19   20  21  22  23  24  day
  допустим сегодня 22 нам нужно найти понедельник для этого мы из 22 - (5-1) = 18
  и ответ проверяем на четность. Если понедельник четное число то неделя четная.
  */
  if(numDayWeek != 0){//номер дня недели не воскресенье 
    var parityWeekMonday = day - (numDayWeek - 1);
    if(parityWeekMonday % 2 == 0){
      parityWeek = ss.getActiveSheet().getRange("A2:H2").getValue();//Указываем клетку с четной неделей 
    }else{
      parityWeek = ss.getActiveSheet().getRange("I2:O2").getValue();//Указываем клетку с нечетной неделей 
    }
  }else{//номер дня недели воскресенье
    var parityWeekMonday = day - 6;
    if(parityWeekMonday % 2 == 0)
    {
      parityWeek = ss.getActiveSheet().getRange("A2:H2").getValue();//Указываем клетку с четной неделей 
    }else{
      parityWeek = ss.getActiveSheet().getRange("I2:O2").getValue();//Указываем клетку с нечетной неделей 
    }
  }
  
  
 /*Проверяем какая сейчас неделя
  с помощью переменной найденной ранее
  если четная то проходимся по массиву дней недели из таблицы и сравниваем с сегодняшним днем недели
  Дальше запоминаем в переменную диапазон из букв нужного нам дня недели (наш столбец)
  и выделяем нужный нам столбец
  если неделя нечетная делаем тоже самое только с массивом букв нечетной недели.
 */ 
  if(String(parityWeek) == "Четная"){
    //ss.getActiveSheet().getRange("F1").setValue("Четная");
    for(var i = 0; i < 7; i++){ 
      if(String(date1[0][i]) == String(weekDayArr[numDayWeek])){
        //ss.getActiveSheet().getRange("G1").setValue(String(date1[0][i]));
        //var dayResDiap = String(dayWeek1[i]) + "1:" + String(dayWeek1[i]);//переменная чтобы выделить столбик
        var dayResDiap = String(dayWeek1[i]) + ":" + String(dayWeek1[i]);//переменная чтобы выделить столбик
        ss.getActiveSheet().getRange(dayResDiap).activate();//выделяем нужный нам столбец
        //ss.getActiveSheet().getRange("G1").setValue(dayResDiap);
      }
    }
  }else{
    //ss.getActiveSheet().getRange("F1").setValue("Не четная");
    for(var i = 0; i < 7; i++){ 
      if(String(date1[0][i]) == String(weekDayArr[numDayWeek])){
        //ss.getActiveSheet().getRange("G1").setValue(String(date1[0][i]));
        //var dayResDiap = String(dayWeek2[i]) + "1:" + String(dayWeek2[i]);//переменная чтобы выделить столбик
        var dayResDiap = String(dayWeek2[i]) + ":" + String(dayWeek2[i]);//переменная чтобы выделить столбик
        ss.getActiveSheet().getRange(dayResDiap).activate();
        //ss.getActiveSheet().getRange("G1").setValue(dayResDiap);
      }
    }
  }

}

