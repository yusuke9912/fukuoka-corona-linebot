const channelToken = "Fd6ReUTr0C9FddEl7uNZlyHrU+hLfEOBb9J4EdASnnTvY1gNhv4N6oSn3canyttuR/E37lBmhKnGJ7XyBKWRCQqY/8ZppdY81P21bO8FJJVreMaVtbZnDagYe63wJq/Yx8ddUSmYjVOeFcvriA2NhwdB04t89/1O/w1cDnyilFU=";
const spreadsheet = SpreadsheetApp.openById("1fq65Ef5Wd6WpTWqng8pXylxGzzG9043ch93M71CNZj0");
const sheet = spreadsheet.getActiveSheet();

const day_list = ["日","月","火","水","木","金","土"];
const regExp = /(?<=（).*?(?=\))/; //()とその中身

//通知を送信
function broadcast() {

sheet.getRange("C11").setValue(Math.random()); //スプレッドシートをリロード

let matches = regExp.exec(sheet.getRange("B3").getValue());
let timestamp = matches[0]; //通知する日付
let notified_date = sheet.getRange("C10").getValue(); //通知済み日付

if(timestamp != notified_date && timestamp != ""){ //送信済みでないことを確認
//データの移動
sheet.getRange("H4:K11").copyTo(sheet.getRange("H3:K10")); 
sheet.getRange("H11").setValue(timestamp);
Logger.log(String(sheet.getRange("C4").getValue()).replace(/\(\d*\)/,"").replace(/※[0-9]/,""));
sheet.getRange("I11").setValue(String(sheet.getRange("C4").getValue()).replace(/\(\d*\)/,"").replace(/※[0-9]/,""));
sheet.getRange("K11").setValue(String(sheet.getRange("C3").getValue()).replace(/\(\d*\)/,"").replace(/※[0-9]/,""));


let today = new Date("2022/"+timestamp);
let day = day_list[today.getDay()];
let title = timestamp + "("+ day +")の感染者数" ;  

let infected_no = String(sheet.getRange("K11").getValue()); //感染者数
let average7 = String(sheet.getRange("C7").getValue()); //直近1週間平均
let average14 = String(sheet.getRange("C8").getValue()); //7日前1週間平均
let diff = sheet.getRange("C9").getValue(); //先週比
let compareAvg = diff <= 0 ? String(diff) : "+" + String(diff); //先週の1週間平均と比較(0より大きい場合は先頭に「+」をつける)


UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", {

method: "post",
headers: {
  "Content-Type": "application/json",
 "Authorization": "Bearer " + channelToken,
},
payload: JSON.stringify({
  messages: [
    {
       "type": "flex",　
    "altText": timestamp + "("+ day +")の感染者数は" + infected_no + "人です",
    "contents" : {
  "type": "bubble",
  "body": {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": title,
            "align": "start",
            "size": "20px"
          },
          {
            "type": "text",
            "align": "start",
            "contents": [
              {
                "type": "span",
                "text": infected_no,
                "color": "#ff4444",
                "weight": "bold",
                "size": "35px"
              },
              {
                "type": "span",
                "text": "人",
                "size": "20px",
                "weight": "regular"
              }
            ],
            "wrap": true
          }
        ],
        "alignItems": "center"
      },
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "直近1週間平均",
            "align": "start",
            "size": "20px"
          },
          {
            "type": "text",
            "align": "start",
            "weight": "bold",
            "contents": [
              {
                "type": "span",
                "text": average7,
                "weight": "bold",
                "size": "35px"
              },
              {
                "type": "span",
                "text": "人",
                "size": "20px",
                "weight": "regular"
              }
            ],
            "wrap": true
          }
        ],
        "alignItems": "center",
        "margin": "5px"
      },
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "7日前1週間平均",
            "align": "start",
            "size": "20px"
          },
          {
            "type": "text",
            "align": "start",
            "contents": [
              {
                "type": "span",
                "text": average14,
                "size": "35px",
                "weight": "bold"
              },
              {
                "type": "span",
                "text": "人",
                "size": "20px",
                "weight": "regular"
              }
            ],
            "wrap": true
          }
        ],
        "alignItems": "center",
        "margin": "5px"
      },
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": "先週比",
            "align": "start",
            "size": "20px"
          },
          {
            "type": "text",
            "align": "start",
            "contents": [
              {
                "type": "span",
                "text": compareAvg,
                "size": "35px",
                "weight": "bold"
              },
              {
                "type": "span",
                "text": "人",
                "size": "20px",
                "weight": "regular"
              }
            ],
            "wrap": true
          }
        ],
        "alignItems": "center",
        "margin": "5px"
      },
    ]
  }
}
    },
  ]
}),
});
prevDate = sheet.getRange("C10").setValue(timestamp); //今日の日付を送信済みとする
}
}
