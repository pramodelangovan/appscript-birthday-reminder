/*
0 Timestamp
1 Email Address
2 Name
3 Birthday
4 Receive Reminders 
*/

function SendBirthdayEmailAndReminder() {
    var date = new Date();
    var day_week = date.getDay();
    var year = date.getFullYear();
    // setting a limit of 7 days
    var limit = 7;
  
    // Creating start time frame and end time frame
    var prefixed_time = ", "+ String(year) + " 03:00:00 UTC";
    var prefixed_end_time = ", "+ String(year) + " 15:00:00 UTC";
    
    // Creating start data and end date
    var start_date = new Date(date.setDate(date.getDate()-day_week))
    var end_date = new Date(date.setDate(date.getDate()+limit))
  
    // Getting start month and end month if a week falls end and starting ofg months
    var start_date_month = Utilities.formatDate(start_date, "IST", "dd");
    var end_date_month = Utilities.formatDate(end_date, "IST", "dd");
  
    // getting month numbers for the calculated start and end dates
    var start_month_number = start_date.getMonth();
    var end_month_number = end_date.getMonth();
  
    // Pushing Months to a month_list array
    var month_list = [];
    if(start_month_number==end_month_number){
      month_list.push(start_month_number)
    } else {
      month_list.push(start_month_number)
      month_list.push(end_month_number)
    }
    
    // contains reminders dates
    var reminder_list = [];
  
    // row starts from 1, we are starting from 2 since 1 contains heading
    var startRow = 2;
  
    // Get the sheet with name reminders
    var birthday = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reminders");
    
    // Getting data matrix dynamically with start and last row and columns
    var birthday_dataRange = birthday.getRange(startRow, 1, birthday.getLastRow(), birthday.getLastColumn());
    
    // Getting json values of rows and column
    var birthday_data = birthday_dataRange.getValues();
    
    var matter = '';
    for(j in month_list){
      for (i in birthday_data) {
        var row = birthday_data[i];
  
        var bDay = new Date(row[3])
        var bDate = Utilities.formatDate(bDay, "IST", "dd");
        var bMonth = bDay.getMonth();
  
        if(bMonth === month_list[j]){
          if(start_month_number==end_month_number){
            if(bDate>=start_date_month && bDate<=end_date_month) {
              matter = matter + birthdayText(row);
              reminder_list.push(birthdayReminder(row));
            }
          } else if(start_month_number==month_list[j]) {
            if(bDate>=start_date_month) {
              matter = matter + birthdayText(row);
              reminder_list.push(birthdayReminder(row));
            }
          } else if(end_month_number==month_list[j]) {
            if(bDate<=end_date_month) {
              matter = matter + birthdayText(row);
              reminder_list.push(birthdayReminder(row));
            }
          }
        }
      }
    }

    if(matter===''){
      matter = "Dear Friends,\n"+
        "There is no Birthday(s) this week\n"+
          "\n"+
            "-\n"+
              "Regards,\n"+
                "Reminder Bot\n"+
                  "\nPlease Note: This is an automated mail, Please do not reply";
  
    } else {
      matter = "Dear Friends,\n"+
        "These are the Birthday(s) for this week\n\n"+
          matter+
          "\nPlease wish them without fail\n"+
            "-\n"+
              "Regards,\n"+
                "Reminder Bot\n"+
                  "\nPlease Note: This is an automated mail, do not reply";
    }
    
    for(k in birthday_data){
      var row = birthday_data[k];
      var res = row[1];
      var subject = 'Birthday(s) Reminder';
      var body = matter
      if(birthday_data[k]!==''){
        if (birthday_data[k][4]==="Yes"){
          GmailApp.sendEmail(res, subject, body);
        }
      }
      
    }
    
    for(e in reminder_list){
      var cal = CalendarApp.getDefaultCalendar();
      var event_start_date = new Date(String(reminder_list[e]['date']) + prefixed_time);
      var event_end_date = new Date(String(reminder_list[e]['date']) + prefixed_end_time);
      var event_message = reminder_list[e]['message'];
      event = cal.createEvent(event_message, event_start_date, event_end_date).addPopupReminder(5).addSmsReminder(5);
      for(k in birthday_data){
        if(birthday_data[k][1]!==''){
          if (birthday_data[k][4]==="Yes"){
            event.addGuest(birthday_data[k][1]);
          }
        }
      }
    }
  }
  
  function birthdayText(row){
    var bDay = row[3]
    var bDate = Utilities.formatDate(bDay, "IST", "dd");
    var bMonth = getMonths()[bDay.getMonth()-1]
    return row[2]+"'s Birthday is on "+bDate+" "+bMonth+"\n"
  }

  function birthdayReminder(row){
    var bDay = row[3]
    var bDate = Utilities.formatDate(bDay, "IST", "dd");
    var bMonth = getMonths()[bDay.getMonth()-1]
    return {'message': row[2]+"'s Birthday", 'date': String(bDate) + " " + String(bMonth) + ","}
  }
  
  function getMonths(){
    // initilizing the months text to list to refer back later
    return ["January","February","March","April","May","June","July","August","September","October","November","December"];
  }