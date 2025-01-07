function doPost(e) {
  try {
    var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
    var userMessage = JSON.parse(e.postData.contents).events[0].message.text;

    Logger.log('User Message: ' + userMessage);

    // 查詢課程資訊
    var courseInfo = searchInSheet(userMessage);
    var replyMessage = '';
    var teacherImageUrl = ''; // 老師圖片 URL，預設為空
    var classroomImageUrl = ''; // 教室圖片 URL，預設為空
    var campusMapUrl = ''; // 校園地圖圖片 URL，預設為空

    if (courseInfo) {
      if (Array.isArray(courseInfo)) {
        // 如果是多門課程（陣列）
        replyMessage = '以下是老師的所有課程資訊：\n\n';
        courseInfo.forEach(function(course) {
          replyMessage += '課程名稱：' + course.courseName + '\n時間：' + course.time + '\n教室：' + course.classroom + '\n學分：' + course.credit + '\n人數：' + course.size + '\n年級：' + course.grade + '\n---\n';
        });

        // 如果查詢到的是老師名稱，只傳老師照片
        if (courseInfo.some(course => course.professor === userMessage)) {
          if (userMessage === '陳建旭') {
            teacherImageUrl = 'https://drive.google.com/uc?id=1j27-gMdOcyjHS0VAyTWmaBvjFSdg62EZ';
          } else if (userMessage === '張婉鈴') {
            teacherImageUrl = 'https://drive.google.com/uc?id=1dwXCfdkG0zSRwAG9cLxCkTfi3qV1Aho-';
          }
          classroomImageUrl = ''; // 確保不回傳教室圖片
          campusMapUrl = ''; // 確保不回傳校園地圖
        }
      } else if (courseInfo.courseName) {
        // 如果是單一課程（物件）
        replyMessage = '課程名稱：' + courseInfo.courseName + '\n老師：' + courseInfo.professor + '\n時間：' + courseInfo.time + '\n教室：' + courseInfo.classroom + '\n學分：' + courseInfo.credit + '\n人數：' + courseInfo.size + '\n年級：' + courseInfo.grade;

        // 如果查詢的是課程名稱，附加教室和校園地圖
        if (['參與式設計', '同步設計', '同步設計研究', '人機互動'].includes(courseInfo.courseName)) {
          classroomImageUrl = 'https://drive.google.com/uc?id=1UDIdUAwOEYldpDmkoor1J1Cj9R7fXhPg';
          campusMapUrl = 'https://drive.google.com/uc?id=1MJFBNkCHOC4SXY2RQdJXmszaKLZQBI3Q';
        }
      } else if (courseInfo.classroom) {
        // 如果是教室號碼
        if (courseInfo.classroom === '5254') {
          classroomImageUrl = 'https://drive.google.com/uc?id=1UDIdUAwOEYldpDmkoor1J1Cj9R7fXhPg';
          campusMapUrl = 'https://drive.google.com/uc?id=1MJFBNkCHOC4SXY2RQdJXmszaKLZQBI3Q';
        } else if (courseInfo.classroom === '5256(B)') {
          classroomImageUrl = 'https://drive.google.com/uc?id=1jp_SJhx5rX--ynbOUGqAL89POaPx7XXD';
          campusMapUrl = 'https://drive.google.com/uc?id=1MJFBNkCHOC4SXY2RQdJXmszaKLZQBI3Q';
        }
      }
    }

    Logger.log('Reply Message: ' + replyMessage);

    // 發送回覆訊息，包含文字與多張圖片
    sendLineMessage(replyToken, replyMessage, teacherImageUrl, classroomImageUrl, campusMapUrl);

  } catch (error) {
    Logger.log('Error in doPost: ' + error.message);
  }
}

function searchInSheet(query) {
  try {
    var sheet = SpreadsheetApp.openById('13PnxBcJM6RylDu7cBRb6F7eNVEj4UOyPYdMLyAhnHb8').getSheetByName('信用卡回饋資訊');
    if (!sheet) {
      Logger.log('Error: Cannot find sheet with name "信用卡回饋資訊"');
      return null;
    }

    var data = sheet.getDataRange().getValues();
    Logger.log('Sheet data: ' + JSON.stringify(data));

    var results = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // 匹配老師、課程名稱或教室號碼
      if (row[0] && typeof row[0] === 'string' && row[0].toLowerCase().trim() === query.toLowerCase().trim()) {
        results.push({
          courseName: row[1],
          professor: row[0],
          time: row[2],
          classroom: row[3],
          credit: row[4],
          size: row[5],
          grade: row[6]
        });
      } else if (row[1] && typeof row[1] === 'string' && row[1].toLowerCase().trim() === query.toLowerCase().trim()) {
        return {
          courseName: row[1],
          professor: row[0],
          time: row[2],
          classroom: row[3],
          credit: row[4],
          size: row[5],
          grade: row[6]
        };
      } else if (row[3] && typeof row[3] === 'string' && row[3].toLowerCase().trim() === query.toLowerCase().trim()) {
        return {
          classroom: row[3]
        };
      }
    }

    if (results.length > 0) {
      return results;
    }

    Logger.log('No Match Found for Query: ' + query);
    return null;

  } catch (error) {
    Logger.log('Error in searchInSheet: ' + error.message);
    return null;
  }
}

function sendLineMessage(replyToken, message, teacherImageUrl, classroomImageUrl, campusMapUrl) {
  try {
    var url = 'https://api.line.me/v2/bot/message/reply';
    var headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + 'oQ7XM8yosJO2lxOxkC6mpeVtt25kuttIymq/ZctH3wgeCFiY7Uctq5ZRaMHIsMN6O/fFE57IYiX/4YJD3GJlAckruWjQbYQ7aEGlFmY14UdeGVO3/oFEXsw8ja5JF502HnNznLjY9v0MLJ0AIBgtVgdB04t89/1O/w1cDnyilFU='
    };

    var messages = [];

    // 添加文字訊息（如果有）
    if (message) {
      messages.push({
        type: 'text',
        text: message
      });
    }

    // 添加老師的圖片訊息
    if (teacherImageUrl) {
      messages.push({
        type: 'image',
        originalContentUrl: teacherImageUrl,
        previewImageUrl: teacherImageUrl
      });
    }

    // 添加教室圖片
    if (classroomImageUrl) {
      messages.push({
        type: 'image',
        originalContentUrl: classroomImageUrl,
        previewImageUrl: classroomImageUrl
      });
    }

    // 添加校園地圖
    if (campusMapUrl) {
      messages.push({
        type: 'image',
        originalContentUrl: campusMapUrl,
        previewImageUrl: campusMapUrl
      });
    }

    var postData = {
      replyToken: replyToken,
      messages: messages
    };

    var options = {
      method: 'post',
      headers: headers,
      payload: JSON.stringify(postData)
    };

    UrlFetchApp.fetch(url, options);

  } catch (error) {
    Logger.log('Error in sendLineMessage: ' + error.message);
  }
}
