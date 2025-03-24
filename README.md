# Extend Closing Automation

## What’s This?
This is a Google Apps Script thing I threw together to handle extend closing requests for container yards when units show up late. It grabs data from a spreadsheet, checks it against vessel times, and fires off emails to the terminal folks with an Excel file attached. Plus, there’s a sidebar UI (`subsequent-extend-closing-UI.html`) for when you need to send more requests manually.

## What It Does
- **1st Sending**: Runs when a vessel’s closing time hits, sending emails for late stuff.
- **Follow-Ups**: Lets you send more requests later with the UI or whatever.
- **Checks Dates**: Makes sure requests fit between vessel open and close times—rejects the rest.
- **Emails**: Sends custom emails depending on the terminal (JICT, KOJA, MAL) with a table and Excel.
- **Triggers**: Sets itself up to run based on closing times from the "Vessel" sheet.
- **UI**: Pick a vessel, pick a sending number, hit send—done.

## Some Code Bits
### Main Loop
```javascript
function processExtendRequests() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Extend Data");
  rows.forEach((row, index) => {
    if (requestTime > opening && requestTime < closing) {
      requestsByVessel[vessel].requests.push({ bookingNumber, containerNo, type, pod, weight, gateIn });
      sheet.getRange(index + 2, 13).setValue("1st Sending");
    }
  });
  for (const vessel in requestsByVessel) {
    sendExtendClosingEmail(vessel, terminal, requests, 1);
  }
}
```

### Email Part
```javascript
function sendExtendClosingEmail(vessel, terminal, newRequests, sendingNumber) {
  const subject = `(${sendingNumber}${ordinalSuffix.toUpperCase()} SENDING) REQUEST EXTEND CLOSING - ${vessel}`;
  const excelBlob = makeExcel(vessel, allRequests); // not the real function, you get it
  MailApp.sendEmail({
    to: "dimasalif5@gmail.com",
    subject: subject,
    htmlBody: htmlBody,
    attachments: [excelBlob]
  });
}
```

### UI Thing
```html
<select id="vesselSelect" onchange="updateSending()"></select>
<button onclick="sendEmail()">Send Whatever’s Left</button>
<script>
  google.script.run.getVessels().then(vessels => {
    const select = document.getElementById('vesselSelect');
    vessels.forEach(v => select.innerHTML += `<option value="${v}">${v}</option>`);
  });
</script>
```

## Why Bother?
If you’re stuck dealing with late containers (like at PT Ocean Network Express Indonesia or wherever), this saves you from typing emails all day. It catches stuff on time, cuts the mistakes, and makes terminals happy with nice emails and files. Less hassle, more done.

## How to Use It
- **Sheets**: Need `Extend Data` (requests) and `Vessel` (schedules—name, opening, closing, terminal).
- **Files**: `extend-closing-automation.gs` runs the show, `subsequent-extend-closing-UI.html` is the UI.
- **Setup**: Run `setupTriggers()` once to get it going at closing times.
- **Permissions**: Give it access to sheets, Drive, and email in Apps Script.

## Running It
- **Auto**: Fires off at closing times.
- **Manual**: Use the "Extend Closing Tools" menu or the sidebar to send more.

It’s not fancy, but it works. Keeps late units moving without the headache.
