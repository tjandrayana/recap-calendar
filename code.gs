function export_gcal_to_gsheet(){

  var start_date = "2021-08-21 00:00:00 GMT+7"
  var finish_date = "2021-08-22 00:00:00 GMT+7"

  var position_start_event = "A"
  var position_finish_event = "B"

  var mycal = "tjandrayana.setiawan@tokopedia.com";
  var cal = CalendarApp.getCalendarById(mycal);

  var timeZone = Session.getScriptTimeZone();

  var date_format = "yyyy-MM-dd"

  var events = cal.getEvents(new Date(start_date), new Date(finish_date), {search: ''});

  events = events.filter(function(e){return  e.getMyStatus() == CalendarApp.GuestStatus.YES || e.getMyStatus() == CalendarApp.GuestStatus.OWNER});

  var sheet = SpreadsheetApp.getActiveSheet();

  var header = [["Event Start", "Do", "Event Title",  "Notes" , "Calculated Duration"]]
  var range = sheet.getRange(1,1,1,5);
  range.setValues(header);


  var iter = 0;
  var notes = ""

  fk =  Utilities.formatDate(events[0].getStartTime(), "GMT", "MM/dd/yyyy");

  var mp = {}
  var mpDuration = {}
    
  for (var i=0;i<events.length;i++) {
    var row=i+2;
    var myformula_placeholder = '';
    
    var d =  Utilities.formatDate(events[i].getStartTime(), "GMT", "MM/dd/yyyy");
    var e = d;
    if (i < events.length - 1){
      e =  Utilities.formatDate(events[i+1].getStartTime(), "GMT", "MM/dd/yyyy");
    }

    if (mp[d] == undefined){
      mp[d] =   "- " + events[i].getTitle() + "\n"
    }else{
      mp[d] =  mp[d] + "- " + events[i].getTitle() + "\n"
    }

    var duration = events[i].getEndTime().valueOf() - events[i].getStartTime().valueOf(); // The unit is millisecond
    var hourDiff = parseInt(duration / ( 60 * 1000)) // Turn the duration into hour format

    if (mpDuration[d] == undefined){
      mpDuration[d] = hourDiff 
    }else{
      mpDuration[d] = mpDuration[d] +  hourDiff 
    }
  }

  var counter = 0
  const days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];

  for (const [key, value] of Object.entries(mp)) {
    row = counter + 2

      var d  = new Date(key);

      var details=[[key , "Meeting on " + days[d.getDay()] , value, "", ""]];
      var range=sheet.getRange(row,1,1,5);
      range.setValues(details);
      counter ++
  }

  counter = 0

  for (const [key, value] of Object.entries(mpDuration)) {
    row = counter + 2    
    var dif  = RoundTo( (parseFloat(value)/60 ) , 0.5 )

    var details=[[dif]];
      var range=sheet.getRange(row,5,1,1);
      range.setValues(details);
      counter ++
  }
}

function RoundTo(number, roundto){
  return roundto * Math.round(number/roundto);
}


function onOpen() {
  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);
}

