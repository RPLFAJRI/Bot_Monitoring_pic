function getAll(e, t) {
  return Sheets.Spreadsheets.Values.get(e, t).values;
}

function findRowIndex(e, t, r) {
  if (e) {
    for (var i = getAll(t, r), n = 0; n < i.length; n++) {
      if (e == idList[n][0]) return parseInt(n + 2);
    }
  }
}

function addNew(e, t, r) {
  var i = Sheets.newRowData();
  return (
    (i.values = e),
    Sheets.Spreadsheets.Values.append(i, t, r, {
      valueInputOption: "USER_ENTERED",
    })
  );
}

function updateData(e, t, r) {
  var i = Sheets.newValueRange();
  return (
    (i.values = e),
    Sheets.Spreadsheets.Values.update(i, t, r, {
      valueInputOption: "USER_ENTERED",
    })
  );
}

function deleteRowById(e, t, r, i) {
  var n = findRowIndex(e, t, i);
  Sheets.Spreadsheets.batchUpdate(
    {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: r,
              dimension: "ROWS",
              startIndex: n,
              endIndex: n + 1,
            },
          },
        },
      ],
    },
    t
  );
}

function uploadfile(e, t, r, i) {
  var n = DriveApp.getFolderById(i),
    a = Utilities.base64Decode(e),
    o = Utilities.newBlob(a, t, r);
  return n.createFile(o).getUrl();
}

function sendFileToTelegram(e, t, r, i) {
  const url = `https://api.telegram.org/bot${e}/sendPhoto`;
  try {
    let response = UrlFetchApp.fetch(url, {
      method: "post",
      payload: { photo: t, caption: r, chat_id: i },
    });
    let jsonResponse = JSON.parse(response.getContentText());
    if (!jsonResponse.ok) {
      console.error("Kesalahan saat mengirim pesan:", jsonResponse.description);
    }
  } catch (error) {
    console.error("Kesalahan dalam mengirim file ke Telegram:", error);
  }
}

function isAccessFile(e) {
  let t = DriveApp.getFileById(e).getSharingAccess(),
    r;
  switch (t) {
    case DriveApp.Access.PRIVATE:
      r = "Private";
      break;
    case DriveApp.Access.ANYONE:
      r = "Anyone";
      break;
    case DriveApp.Access.ANYONE_WITH_LINK:
      r = "Anyone with a link";
      break;
    case DriveApp.Access.DOMAIN:
      r = "Anyone inside domain";
      break;
    case DriveApp.Access.DOMAIN_WITH_LINK:
      r = "Anyone inside domain who has the link";
      break;
    default:
      r = "Unknown";
  }
  console.log("File has been permission: ", r);
  return ["Anyone", "Anyone with a link"].indexOf(r) > -1;
}

function exportRangeToFileBlob(e, t) {
  let r = convertFileToPdfUrl(e, t, {
    measureLimit: 150,
    sizeLimit: 1600,
    imageScale: 1,
  });
  let i = UrlFetchApp.fetch(r, {
    muteHttpExceptions: !0,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
  });

  if (200 !== i.getResponseCode()) {
    console.log("Response Code: " + i.getResponseCode());
    return;
  }

  let n = i.getBlob();
  return n;
}

function convertFileToPdfUrl(e, t, r) {
  var i = t.getSheet(),
    n = e.getUrl(),
    a = i.getSheetId(),
    o = t.getRow(),
    s = t.getColumn(),
    g = t.getLastRow(),
    l = t.getLastColumn();

  if (g - o + 1 > r.sizeLimit)
    throw "The range exceeded the limit of " + r.sizeLimit + " rows";
  if (l - s + 1 > r.sizeLimit)
    throw "The range exceeded the limit of " + r.sizeLimit + " columns";

  for (var c = 0, d = 0, u = s; u <= l; u++) {
    u <= r.measureLimit && (d = i.getColumnWidth(u));
    c += d + 0.3;
    //Logger untuk mencatat proses
    if (u % 50 == 0 && u <= r.measureLimit) {
      Logger.log("Selesai " + u + " kolom dari " + l);
    }
  }

  for (var m = 0, h = 0, u = o; u <= g; u++) {
    u <= r.measureLimit && (h = i.getRowHeight(u));
    m += h + 1;
    if (u % 50 == 0 && u <= r.measureLimit) {
      Logger.log("Selesai " + u + " baris dari " + g);
    }
  }

  var E = {
    url: n,
    sheetId: a,
    r1: o - 1,
    r2: g,
    c1: s - 1,
    c2: l, //lebar bawah untuk atur '156'
    size:
      Math.round((c / 100) * 1e3 + 100) / 1e3 +
      "x" +
      Math.round((m / 156) * 1e3 + 100) / 1e3, // Ubah ukuran
    scale: 2, // Ubah skala
    top_margin: 0,
    bottom_margin: 0,
    left_margin: 0,
    right_margin: 0,
  };
  var f = "&r1=" + E.r1 + "&r2=" + E.r2 + "&c1=" + E.c1 + "&c2=" + E.c2,
    p = "&gid=" + E.sheetId,
    $ = "";

  return (
    E.url.replace(/\/edit.*$/, "") +
    "/export?exportFormat=pdf&format=pdf&size=" +
    E.size +
    p +
    f +
    "&scale=" +
    E.scale +
    "&top_margin=" +
    E.top_margin +
    "&bottom_margin=" +
    E.bottom_margin +
    "&left_margin=" +
    E.left_margin +
    "&right_margin=" +
    E.right_margin +
    "&sheetnames=false&printtitle=false&pagenum=UNDEFINED&horizontal_alignment=LEFT&gridlines=false&fmcmd=12&fzr=FALSE"
  );
}

// Fungsi untuk penjadwalan dan penghapusan trigger
const jobType = {
  EVERY_MINUTES: "EVERY_MINUTES",
  EVERY_HOURS: "EVERY_HOURS",
  EVERY_DAYS: "EVERY_DAYS",
  EVERY_WEEKS: "EVERY_WEEKS",
  AT: "AT",
};

function scheduleJob(e, t, r) {
  switch (t) {
    case jobType.AT:
      return ScriptApp.newTrigger(e)
        .timeBased()
        .everyDays(1)
        .atHour(r)
        .inTimezone("GMT+7")
        .create();
    case jobType.EVERY_MINUTES:
      return ScriptApp.newTrigger(e).timeBased().everyMinutes(r).create();
    case jobType.EVERY_HOURS:
      return ScriptApp.newTrigger(e).timeBased().everyHours(r).create();
    case jobType.EVERY_DAYS:
      return ScriptApp.newTrigger(e).timeBased().everyDays(r).create();
    case jobType.EVERY_WEEKS:
      return ScriptApp.newTrigger(e).timeBased().everyWeeks(r).create();
  }
}

function deleteTriggers() {
  for (var e = ScriptApp.getProjectTriggers(), t = 0; t < e.length; t++) {
    ScriptApp.deleteTrigger(e[t]);
  }
}

function deleteTrigger(e) {
  for (var t = ScriptApp.getProjectTriggers(), r = 0; r < t.length; r++) {
    if (t[r].getHandlerFunction() == e) {
      ScriptApp.deleteTrigger(t[r]);
    }
  }
}

function scheduleEveryDayJob(functionName) {
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone("GMT+7")
    .create();
}
