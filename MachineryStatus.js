function O99_NarzedziaPomiarowe() {

  var ss = SpreadsheetApp.openById('******').getSheetByName("O99 - NARZĘDZIA POMIAROWE");
  //otwarcie odpowiedniej karty w pliku  Status Maszyn 

  var loopmax = 160;
  var today = new Date;
  var todayYear = today.getUTCFullYear();
  var todayMonth = today.getUTCMonth() + 1;
  var subject = 'Lista przyrządow do kalibracji w miesiącu:' + todayMonth + '.' + todayYear;
  //utworzenie zmiennej definiującej temat e-maila
  var message = 'Rodzaj przyrządu - Nazwa stanowiska - Komentarz - Odp. za kalibrację \n';
  //utworzenie nagłówka e-maila

  var arrTotal = ss.getRange(3,1,loopmax,12).getDisplayValues();
  var licznikloop = 0;

  for (var i = 0; i < loopmax; ++i) 
  {
    //pętla - UWAGA w przypadku przekroczenia liczby 160 przyrządów należy pętlę odpowiednio wydłużyć
    var a = arrTotal[i][10];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][8] === 'O99' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][0] 
      + ' - ' 
      + arrTotal[i][1]
      + ' - ' 
      + arrTotal[i][5]
      + ' - '
      + arrTotal[i][6]
      + ' - ' 
      + arrTotal[i][11] 
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK PRZYRZĄDÓW DO KALIBRACJI';
  }
  
  message = message + '\n\nLink do raportu:\nhttps://docs.google.com/spreadsheets/d/*******/edit#gid=*******';

  MailApp.sendEmail('marcin.z@x.com', subject, message);
  //MailApp.sendEmail('jolanta@x.com', subject, message);
  MailApp.sendEmail('f@x.com', subject, message);
  MailApp.sendEmail('sabina@x.com', subject, message);
  MailApp.sendEmail('oliwi@x.com', subject, message);

}

function O99_MaszynyIStanowiska() {
  var ss = SpreadsheetApp.openById('*****').getSheetByName("O99 - MASZYNY I STANOWISKA");
  //otwarcie odpowiedniej karty w pliku Status Maszyn 
  var loopmax = 300;
  var today = new Date;
  var todayYear = today.getUTCFullYear();
  var todayMonth = today.getUTCMonth() + 1;
  var nextMonth;
  var nextMonthYear;

  if(todayMonth == 12)
  {
    nextMonth = 1;
    nextMonthYear = todayYear+1;
  }
  else
  {
    nextMonth = todayMonth +1;
    nextMonthYear = todayYear;
  }

  var arrTotal = ss.getRange(3,1,300,29).getDisplayValues();
  var licznikloop = 0;
  var subject = 'Przegląd maszyn i urządzeń na miesiąc:' + todayMonth + '.' + todayYear;
  var message = 
  'KONTROLA ELEKTRYCZNA MASZYN W TYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data testów elektrycznych \n';   

  for (var i = 0; i < loopmax; ++i){
    var a = arrTotal[i][21];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'TB' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][10] === 'TAK' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][21] 
      + '\n';
    }
    if 
    (
      arrTotal[i][1] === 'M' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][10] === 'TAK' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][21] 
      + '\n';
    }
    if 
    (
      arrTotal[i][1] === 'X' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][10] === 'TAK' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][21] 
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

  licznikloop=0;  
  message = message +
  '\nKONTROLA ELEKTRYCZNA MASZYN W PRZYSZŁYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data testów elektrycznych \n';

  for (var i = 0; i < loopmax; ++i){
    var a = arrTotal[i][21];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'TB' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][10] === 'TAK' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
 

    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][21] 
      + '\n';
    }
    if 
    (
      arrTotal[i][1] === 'M' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][10] === 'TAK' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
 

    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][21] 
      + '\n';
    }
    if 
    (
      arrTotal[i][1] === 'X' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][10] === 'TAK' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
 

    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][21] 
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

  licznikloop=0;  
  message = message +
  '\nKONTROLA RCD MASZYN W TYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data kontroli RCD \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][23];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'M' &&
      arrTotal[i][9] === 'O99' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;        
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][23]
      + '\n';
    }
    if 
    (
      arrTotal[i][1] === 'X' &&
      arrTotal[i][9] === 'O99' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;        
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][23]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }
  
  licznikloop=0;  
  message = message +
  '\nKONTROLA RCD MASZYN W PRZYSZŁYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data kontroli RCD \n';

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][23];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'M' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][22] === 'TAK' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][23]
      + '\n';
    }
    if 
    (
      arrTotal[i][1] === 'X' &&
      arrTotal[i][9] === 'O99' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][23]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }
  
  licznikloop=0;  
  message = message +
  '\nKONTROLA TERMOWIZYJNA SZAFY ELEKTRYCZNEJ MASZYN W TYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data kontroli Termowizyjnej \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][25];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if
    (
      arrTotal[i][1] === 'M' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][24] === 'TAK' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][25]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

  licznikloop=0;  
  message = message +
  '\nKONTROLA TERMOWIZYJNA SZAFY ELEKTRYCZNEJ MASZYN W PRZYSZŁYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data kontroli Termowizyjnej \n';

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][25];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'M' &&
      arrTotal[i][9] === 'O99' &&
      arrTotal[i][24] === 'TAK' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][25]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

  licznikloop=0;  
  message = message +
  '\nPRZEGLĄDY STANOWISK I REGAŁÓW W TYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data przeglądu \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][28];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth()+1 ;

    if 
    (
      arrTotal[i][1] === 'R' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;        
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][28]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }
  
  licznikloop=0;  
  message = message +
  '\nPRZEGLĄDY STANOWISK I REGAŁÓW W PRZYSZŁYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data przeglądu \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][28];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'R' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
    {
      licznikloop ++;
      message = message 
      + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][28]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

    licznikloop=0;  
  message = message +
  '\nPRZEGLĄDY ODCIĄGÓW W TYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data przeglądu \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][28];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth()+1 ;

    if 
    (
      arrTotal[i][1] === 'OD' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;        
      message = message 
     // + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][28]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

    licznikloop=0;  
  message = message +
  '\nPRZEGLĄDY ODCIĄGÓW W PRZYSZŁYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data przeglądu \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][28];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'OD' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
    {
      licznikloop ++;
      message = message 
      //+ arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][28]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

      licznikloop=0;  
  message = message +
  '\nPRZEGLĄDY RĘKAWIC ELEKTROIZOLACYJNYCH W TYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data przeglądu \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][28];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth()+1 ;

    if 
    (
      arrTotal[i][1] === 'RE' &&
      expDateMonth === todayMonth &&
      expDateYear === todayYear
    )
    {
      licznikloop ++;        
      message = message 
     // + arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][28]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }

      licznikloop=0;  
  message = message +
  '\nPRZEGLĄDY RĘKAWIC ELEKTROIZOLACYJNYCH W PRZYSZŁYM MIESIĄCU \n' +
  'Oznaczenie stanowiska - Nazwa stanowiska - Data przeglądu \n'; 

  for (var i = 0; i < loopmax; ++i) {
    var a = arrTotal[i][28];
    var b = new Date(a);
    var expDateYear = b.getUTCFullYear();
    var expDateMonth = b.getUTCMonth() + 1;

    if 
    (
      arrTotal[i][1] === 'RE' &&
      expDateMonth === nextMonth &&
      expDateYear === nextMonthYear
    )
    {
      licznikloop ++;
      message = message 
      //+ arrTotal[i][1] 
      + arrTotal[i][2] 
      + ' - ' 
      + arrTotal[i][3] 
      + ' - ' 
      + arrTotal[i][28]
      + '\n';
    }
  }

  if (licznikloop === 0)
  {
    message = message + 'BRAK ZAPLANOWANYCH AKCJI\n';
  }



  message = message + 
  '\nLINK DO PLIKU:\nhttps://docs.google.com/spreadsheets/d/*****/edit#gid=*****';


  MailApp.sendEmail('marcin@x.com', subject, message);
  MailApp.sendEmail('krzysztof@x.com', subject, message);
  MailApp.sendEmail('f@x.com', subject, message);

}