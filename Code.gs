function fetchDewPointData() {

  // Variables
  var recipientEmail = "YOUR_EMAIL_ADDRESS";
  var emailSubject = "Humidity Discomfort - Next 6 Days";
  var latitude = "YOUR_LATITUDE_DECIMAL";
  var longitude = "YOUR_LONGITUDE_DECIMAL";
  var apikey = "YOUR_API_KEY";
  var dewpointDescriptions = {
    'level_1': '#FF0000',  // worst
    'level_2': '#FF4500',  // almost worst
    'level_3': '#FF8C00',  // uncomfortable
    'level_4': '#FFFF00',  // sticky
    'level_5': '#00FF00',  // comfortable
    'level_6': '#32CD32',  // pleasant
    'level_7': '#87CEEB',  // dry
    'level_8': '#0000FF'   // extremely dry
  };
  var dewpointRanges = {
    'level_1': [75, 100],
    'level_2': [70, 74],
    'level_3': [65, 69],
    'level_4': [60, 64],
    'level_5': [55, 59],
    'level_6': [50, 54],
    'level_7': [32, 49],
    'level_8': [0, 31]
  };
  var dewpointRangeName = {
    'level_1': 'max',
    'level_2': 'plus 3',
    'level_3': 'plus 2',
    'level_4': 'plus 1',
    'level_5': 'norm',
    'level_6': 'norm',
    'level_7': 'dry',
    'level_8': 'min'
  };
  var IncreaseColor = "#FF0000";
  var DecreaseColor = "#0000FF";
  var skip_if_all_same = false;

  try {
    var url = 'https://api.tomorrow.io/v4/timelines?location=' + latitude + ',' + longitude + '&fields=dewPoint,temperatureApparent,temperature&timesteps=1d&units=imperial&apikey=' + apikey;
    var response = UrlFetchApp.fetch(url);
    var json = response.getContentText();
    var data = JSON.parse(json);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    sheet.getRange('A1').setValue('Time');
    sheet.getRange('B1').setValue('Temperature');
    sheet.getRange('C1').setValue('Calculated Humidity');
    sheet.getRange('D1').setValue('Dew Point');
    sheet.getRange('E1').setValue('Comfort Category');
    sheet.getRange('F1').setValue('Feels Like Temp Difference');

    var emailBody = '';
    var previousFeelsLikeTemp;
    var comfortCategories = [];

    for (var i = 0; i < data.data.timelines[0].intervals.length; i++) {
      var row = i + 2;
      var timestamp = new Date(data.data.timelines[0].intervals[i].startTime);
      var formattedDate = Utilities.formatDate(timestamp, "America/Chicago", 'EEE dd');
      var dewPointFahrenheit = Math.round(data.data.timelines[0].intervals[i].values.dewPoint); // Rounded Dew Point
      var temperatureFahrenheit = data.data.timelines[0].intervals[i].values.temperature;
      var temperatureFahrenheitRnd = Math.round(data.data.timelines[0].intervals[i].values.temperature); // Rounded Temperature
      var humidity = 100 * Math.exp((17.27 * ((dewPointFahrenheit - 32) * 5/9)) / (237.7 + ((dewPointFahrenheit - 32) * 5/9)) - (17.27 * ((temperatureFahrenheit - 32) * 5/9)) / (237.7 + ((temperatureFahrenheit - 32) * 5/9)));
      var humidityRnd = Math.round(humidity); // Rounded Humidity (not sure if this'll round to 1)
      var feelsLikeTemp = Math.round(data.data.timelines[0].intervals[i].values.temperatureApparent);
      var tempDifference;
      var formattedTempDifference;

      if (i !== 0) {
        tempDifference = feelsLikeTemp - previousFeelsLikeTemp;
		formattedTempDifference = (tempDifference >= 0 ? '+' : '-') + Math.round(Math.abs(tempDifference)).toString().padStart(2, '0');
		formattedTempDifferenceColor = tempDifference > 0 ? IncreaseColor : (tempDifference < 0 ? DecreaseColor : '');
		formattedcoloredTempDifference = '<span style="color: ' + formattedTempDifferenceColor + '">' + formattedTempDifference + '</span>';
      } else {
        formattedcoloredTempDifference = '&nbsp;' + feelsLikeTemp.toString().padStart(2, " ");
      }

      previousFeelsLikeTemp = feelsLikeTemp;

      var comfortCategory;
      var range;
      var rangeName;
      if (dewPointFahrenheit >= 75) {
        comfortCategory = dewpointDescriptions.level_1;
        range = dewpointRanges.level_1;
        rangeName = dewpointRangeName.level_1;
      } else if (dewPointFahrenheit >= 70) {
        comfortCategory = dewpointDescriptions.level_2;
        range = dewpointRanges.level_2;
        rangeName = dewpointRangeName.level_2;
      } else if (dewPointFahrenheit >= 65) {
        comfortCategory = dewpointDescriptions.level_3;
        range = dewpointRanges.level_3;
        rangeName = dewpointRangeName.level_3;
      } else if (dewPointFahrenheit >= 60) {
        comfortCategory = dewpointDescriptions.level_4;
        range = dewpointRanges.level_4;
        rangeName = dewpointRangeName.level_4;
      } else if (dewPointFahrenheit >= 55) {
        comfortCategory = dewpointDescriptions.level_5;
        range = dewpointRanges.level_5;
        rangeName = dewpointRangeName.level_5;
      } else if (dewPointFahrenheit >= 50) {
        comfortCategory = dewpointDescriptions.level_6;
        range = dewpointRanges.level_6;
        rangeName = dewpointRangeName.level_6;
      } else if (dewPointFahrenheit >= 32) {
        comfortCategory = dewpointDescriptions.level_7;
        range = dewpointRanges.level_7;
        rangeName = dewpointRangeName.level_7;
      } else {
        comfortCategory = dewpointDescriptions.level_8;
        range = dewpointRanges.level_8;
        rangeName = dewpointRangeName.level_8;
      }
      comfortCategories.push(comfortCategory);

      sheet.getRange('A' + row).setValue(timestamp);
      sheet.getRange('B' + row).setValue(temperatureFahrenheit);
      sheet.getRange('C' + row).setValue(humidity);
      sheet.getRange('D' + row).setValue(dewPointFahrenheit);
      sheet.getRange('E' + row).setValue(comfortCategory);
      sheet.getRange('E' + row).setBackgroundColor(comfortCategory);
      sheet.getRange('F' + row).setValue(formattedTempDifference + (i !== 0 ? ' (' + feelsLikeTemp + ')' : ''));
  
      // Updated email body
// emailBody += '<div style="background-color:' + comfortCategory + '; display: inline;">&nbsp;&nbsp;&nbsp;' + dewPointFahrenheit + '° w/in ' + range[0] + '-' + range[1] + '° (' + rangeName + ')&nbsp;&nbsp;&nbsp;</div> <b>' + formattedDate + '</b>: feels: ' + feelsLikeTemp + (i !== 0 ? '° (' + formattedcoloredTempDifference + ')' : '°') + ' (dpc: ' + temperatureFahrenheitRnd + '° × ' + humidityRnd + '%)' + '<br>';
emailBody += '<div style="background-color:' + comfortCategory + '; display: inline;">&nbsp;&nbsp;&nbsp;.&nbsp;&nbsp;&nbsp;</div> <b>' + formattedDate + '</b> &ndash; ' + rangeName + ' &ndash; (' + temperatureFahrenheitRnd + '° x ' + humidityRnd + '%) &ndash; ' + 'feels: ' + feelsLikeTemp + (i !== 0 ? '° (' + formattedcoloredTempDifference + ')' : '°') + '<br>';
    }

    if (skip_if_all_same && comfortCategories.every((val, i, arr) => val === arr[0])) {
        return; // Skip sending the email
    }

    MailApp.sendEmail({
      to: recipientEmail,
      subject: emailSubject,
      body: emailBody,
      htmlBody: emailBody
    });
  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}
