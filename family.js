function SendBirthdayEmailAndReminder() {
    // initilizing the months text to list to refer back later
    var months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  
  
    var date = new Date();
    var day_week = date.getDay();
    var year = date.getFullYear();
    var limit = 7;
    var prefixed_time = ", "+ String(year) + " 03:00:00 UTC";
    var prefixed_end_time = ", "+ String(year) + " 15:00:00 UTC";
    var start_date = new Date(date.setDate(date.getDate()-day_week))
    var end_date = new Date(date.setDate(date.getDate()+limit))
  
    var start_date_month = Utilities.formatDate(start_date, "IST", "dd");
    var end_date_month = Utilities.formatDate(end_date, "IST", "dd");
  
    var start_month_number = start_date.getMonth();
    var end_month_number = end_date.getMonth();
  
    var month_list = [];
    if(start_month_number==end_month_number){
      month_list.push(start_month_number)
    } else {
      month_list.push(start_month_number)
      month_list.push(end_month_number)
    }
    
    var reminder_list = [];
  
    var startRow = 2;
    var birthday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Birthday");
    var annivesary = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Anniversary");
    var email = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email");
  
    var birthday_dataRange = birthday.getRange(startRow, 1, birthday.getLastRow(), birthday.getLastColumn());
    var annivesary_dataRange = annivesary.getRange(startRow, 1, annivesary.getLastRow(), annivesary.getLastColumn());
    var email_dataRange = email.getRange(1, 1, email.getLastRow(), email.getLastColumn());
      
    var birthday_data = birthday_dataRange.getValues();
    var annivesary_data = annivesary_dataRange.getValues();
    var email_data = email_dataRange.getValues();
   
  
    var matter = '';
    for(j in month_list){
      for (i in birthday_data) {
        var row = birthday_data[i];
        if(row[3] === months[month_list[j]]){
          if(start_month_number==end_month_number){
            if(row[2]>=start_date_month && row[2]<=end_date_month) {
              matter = matter + birthdayText(row);
              reminder_list.push(birthdayReminder(row));
            }
          } else if(start_month_number==month_list[j]) {
            if(row[2]>=start_date_month) {
              matter = matter + birthdayText(row);
              reminder_list.push(birthdayReminder(row));
            }
          } else if(end_month_number==month_list[j]) {
            if(row[2]<=end_date_month) {
              matter = matter + birthdayText(row);
              reminder_list.push(birthdayReminder(row));
            }
          }
        }
      }
  
      for (i in annivesary_data) {
        var row = annivesary_data[i];
        if(row[4] === months[month_list[j]]){
          if(start_month_number==end_month_number){
            if(row[3]>=start_date_month && row[3]<=end_date_month) { 
              matter = matter + annivesaryText(row);
              reminder_list.push(annivesaryReminder(row))
            }
          } else if(start_month_number==month_list[j]) {
            if(row[3]>=start_date_month) { 
              matter = matter + annivesaryText(row);
              reminder_list.push(annivesaryReminder(row));
            }
          } else if(end_month_number==month_list[j]) {
            if(row[3]<=end_date_month) { 
              matter = matter + annivesaryText(row);
              reminder_list.push(annivesaryReminder(row));
            }
          }
        }
      }
  
    }
    if(matter===''){
      matter = "Hi Family,\n"+
        "There is no Birthday or Anniversaries this week\n"+
          "\n"+
            "-\n"+
              "Regards,\n"+
                "Pramod\n"+
                  "\nPlease Note: This is an automated mail, Please do not reply";
  
    } else {
      matter = "Hi Family,\n"+
        "These are the Birthday and Anniversaries for this week\n\n"+
          matter+
          "\nPlease wish them without fail\n"+
            "-\n"+
              "Regards,\n"+
                "Reminder Bot\n"+
                  "\nPlease Note: This is an automated mail, do not reply";
    }
    
    for(k in email_data){
      var row = email_data[k];
      var res = row[0];
      var subject = 'Birthday and Anniversaries Reminder';
      var body = matter
      if(email_data[k]!==''){
        GmailApp.sendEmail(res, subject, body);
      }
      
    }
    
    for(e in reminder_list){
      var cal = CalendarApp.getDefaultCalendar();
      var event_start_date = new Date(String(reminder_list[e]['date']) + prefixed_time);
      var event_end_date = new Date(String(reminder_list[e]['date']) + prefixed_end_time);
      var event_message = reminder_list[e]['message'];
      event = cal.createEvent(event_message, event_start_date, event_end_date).addPopupReminder(5).addSmsReminder(5);
      for(k in email_data){
        if(email_data[k][0]!==''){
          event.addGuest(email_data[k][0]);
        }
      }
    }
  }
  
  function birthdayText(row){
    return row[0]+"'s "+row[1]+" is on "+row[2]+" "+row[3]+"\n"
  }
  
  function annivesaryText(row){
    return row[0]+" and "+row[1]+" are having thier "+ row[2] +" on "+row[3]+" "+row[4]+"\n"
  }
  
  function birthdayReminder(row){
    return {'message': row[0]+"'s "+row[1], 'date': String(row[3]) + " " + String(row[2]) + ","}
  }
  
  function annivesaryReminder(row){
    return {'message' : row[0]+" and "+row[1]+" are having thier "+ row[2], 'date' : String(row[4]) + " " + String(row[3]) + ","}
  }
  
  function makeReminders() {
    var cal = CalendarApp.getDefaultCalendar();
  
    event = cal.createEvent('Today is someone\'s birthday', new Date('February 14, 2020 11:50:00 UTC'), new Date('February 14, 2020 12:00:00 UTC')).addPopupReminder(1).addSmsReminder(1);
    event.addGuest('pramood46@gmail.com');
  }
  