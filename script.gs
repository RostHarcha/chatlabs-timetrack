function msToTime(duration) {
    var milliseconds = Math.floor((duration % 1000) / 100),
        seconds = Math.floor((duration / 1000) % 60),
        minutes = Math.floor((duration / (1000 * 60)) % 60),
        hours = Math.floor((duration / (1000 * 60 * 60)) % 24);
    hours = (hours < 10) ? "0" + hours : hours;
    minutes = (minutes < 10) ? "0" + minutes : minutes;
    return hours + ":" + minutes;
}

function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

function getWorkTime(date, projectID, cookie, users) {
    const fromDate = date.valueOf();
    const toDate = addDays(date, 1).valueOf();
    let apiUrl = 'https://app.pendulums.io/api/projects';
    const res = UrlFetchApp.fetch(
        apiUrl + '/' + projectID + '/stats/hours?users=' + users + '&from=' + fromDate + '&to=' + toDate,
        {
            'method': 'get',
            'contentType': 'application/json',
            'headers': {
                'cookie': cookie
            }
        }
    );
    const milliseconds = JSON.parse(res.getContentText()).result[0].stats[0].value;
    return msToTime(milliseconds);
}

function getRaw(sheet, raw, start_column) {
    while (sheet.getRange(raw, start_column).getValue()) {
        raw++;
    }
    return raw;
}

function formatDate(date) {
    var dd = String(date.getDate()).padStart(2, '0');
    var mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
    var yyyy = date.getFullYear();
    return dd + '.' + mm + '.' + yyyy;
}

function getSheet(url, name) {
    const SPREADSHEET_URL = url;
    const SHEET_NAME = name;
    const ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    return ss.getSheetByName(SHEET_NAME);
}

function convertDate(date) {
    const split_date = date.split('.');
    return new Date(split_date[2], split_date[1] - 1, split_date[0], -7);
}

function addWork() {
    const sheetUrl = 'https://docs.google.com/spreadsheets/d/1ytwfBvslx5Zs9RTKHJBklhaowQ4iGfYGfQGYjrIn-ko/';
    const sheetName = 'ChatLabs';
    const sheet = getSheet(sheetUrl, sheetName);
    const date = Utilities.formatDate(sheet.getRange('F21').getValue(), 'Europe/Moscow', 'dd.MM.yyyy');
    const datetime = convertDate(date);
    const cookie = sheet.getRange('F22').getValue();
    const users = sheet.getRange('F23').getValue();
    const projects = sheet.getRange('K3:L19').getValues();
    for (let raw in projects) {
        let projectName = projects[raw][0];
        let projectID = projects[raw][1];
        if (projectID == '') {
            continue;
        }
        let workTime = getWorkTime(datetime, projectID, cookie, users);
        if (workTime == '00:00') {
            continue;
        }
        sheet.getRange(getRaw(sheet, 3, 2), 2, 1, 3).setValues([[
            date,
            workTime,
            projectName
        ]]);
    }
}
