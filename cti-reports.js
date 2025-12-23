function sendEmail2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var notify_staff = sheet.getRange("J:J").getValues();
  var patient_name = sheet.getRange("A:A").getValues();
  var patient_cti_date = sheet.getRange("C:C").getValues();
  var patient_notify_date = sheet.getRange("F:F").getValues();
  const email_list = ["kobet@parrishhealthsystems.org","reneesha@parrishhealthsystems.org"];
  const weekday = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const d = new Date();

  function isDateWithinRange(targetDate, startDate, endDate) {
  return targetDate >= startDate && targetDate <= endDate;
  }
  
  //notify_staff.forEach(function(){
   // Logger.log(notify_staff)
 // })
  
  let day = weekday[d.getDay()];
  let weekly_summary_email = [];
  if (day == "Thursday") {
  const vs = patient_notify_date.flat();
  const wk = parseInt(Utilities.formatDate(new Date(), "GMT", "w"));
  vs.forEach((d,i) => {
    let w =parseInt(Utilities.formatDate(new Date(d), "GMT", "w"));
    if(w == wk || w == wk - 1) {//try this for last two weeks. This may not work at beginning of year...I don't know for sure.
      Logger.log(patient_name[i])
      weekly_summary_email.push(patient_name[i] + " - Current Period [" + patient_cti_date[i] + "]\n")
    } else (
      Logger.log("")
    )
  });

  email_list.forEach((email,v) => {
      if (day == "Thursday") {
      GmailApp.sendEmail(
        email,
        '[CTI Notification System] Weekly Summary',
        "Patients with upcoming dates this month.\n" + weekly_summary_email,
      )
    }
  })
  }
}
