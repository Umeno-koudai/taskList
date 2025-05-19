/*命名規則
定数 = アッパースネークケース(START_CLOUMN)
変数 = スネークケース(current_row)
関数 = キャメルケース(getUserName)
コンポーネント = アッパーキャメル(UserFrom)
プロパティ = ローワーキャメル(userName)
クラス = アッパーキャメル(MayClass)*/

//アクティブなシートを取得
const ACTIVE_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
//アクティブなシートを取得
const ACTIVE_SHEET = ACTIVE_SPREADSHEET.getActiveSheet();
var UI;

const TASK_NAME = 2;    //タスク名の列番号
const START_DAY = 3;    //開始日の列番号
const END_DAY = 4;      //納期の列番号
const START_TIME = 5;   //開始時間の列番号
const END_TIME = 6;     //終了時間の列番号
const IS_COLENDER = 8;  //カレンダーに登録ずみか確認する列番号
const CALEMDER_TYPE = 9;

const START_ROW = 3;    //表のスタート位置の行番号
const START_COLUMN = 1; //表のスタート位置の列番号
const LAST_COLUMN = 9;  //表の最後の列番号

function onOpen(){
  UI = SpreadsheetApp.getUi();
}

//タスクデータを格納するクラス
class task{
  constructor(taskName, startDate, endDate, startTime, endTime, calenderType){
    this.taskName = taskName;
    this.startDay = startDate;
    this.endDay = endDate;
    this.startTime = startTime;
    this.endTime = endTime;
    this.colenderType = calenderType;
  }
}

//タスク更新メソッド(カレンダーに登録していないタスクをカレンダーに登録する)
function taskUpdate(){
  //一番最後の行を取得
  let last_row = ACTIVE_SHEET.getLastRow();
  //タスク表を全て取得
  const content = ACTIVE_SHEET.getRange(START_ROW, START_COLUMN, last_row-2, LAST_COLUMN).getValues();

  let calender = CalendarApp.getDefaultCalendar();

  //取得したタスク表をもとにカレンダーにタスクを登録していく
  for(let i = 0; i < content.length; i++){
    if(content[i][0] === false && content[i][IS_COLENDER-1] === false){
      let current_data = new task(content[i][TASK_NAME-1],content[i][START_DAY-1],content[i][END_DAY-1],content[i][START_TIME-1],content[i][END_TIME-1], content[i][CALEMDER_TYPE-1]);
      console.log(current_data);

      if(current_data.colenderType !== ""){
        let calenders = CalendarApp.getCalendarsByName(current_data.colenderType);
        if(calenders.length > 0){
          calender = calenders[0];
        }
        else{
          calender = CalendarApp.createCalendar(current_data.colenderType);
        }
      }

      if(current_data.startDay){
        let startDate = new Date(current_data.startDay);
        let endDate;
        if(current_data.endDay){
          endDate = new Date(current_data.endDay);
        }
        else{
          endDate = new Date(current_data.startDay);
          ACTIVE_SHEET.getRange(START_ROW+i, END_DAY).setValue(current_data.startDay);
        }

        if(current_data.startTime && current_data.endTime){
          startDate.setHours(current_data.startTime.getHours());
          startDate.setMinutes(current_data.startTime.getMinutes());

          endDate.setHours(current_data.endTime.getHours());
          endDate.setMinutes(current_data.endTime.getHours());
          calender.createEvent(current_data.taskName, startDate, endDate);
        }
        else{
          endDate.setDate(endDate.getDate() + 1);
          calender.createAllDayEvent(current_data.taskName, startDate, endDate);
        }
      }
      else{
        //UI.alert("日付が入力されていません");
      }
      ACTIVE_SHEET.getRange(START_ROW+i, IS_COLENDER).setValue(true);
      console.log(ACTIVE_SHEET.getRange(START_ROW+i, IS_COLENDER).getValue());
      
    }
  }
}

//新しいタスク行を追加する
function addTask(){
  let last_row = ACTIVE_SHEET.getLastRow();
  let data = ACTIVE_SHEET.getRange(last_row, START_COLUMN, 1, LAST_COLUMN);
  data.copyTo(ACTIVE_SHEET.getRange(last_row+1, START_COLUMN, 1, LAST_COLUMN), SpreadsheetApp.CopyPasteType.PASTE_FORMAT);
  ACTIVE_SHEET.getRange(last_row+1, 1).setValue(false);
  ACTIVE_SHEET.getRange(last_row+1, TASK_NAME).setValue("");
  ACTIVE_SHEET.getRange(last_row+1, START_DAY).setValue("");
  ACTIVE_SHEET.getRange(last_row+1, END_DAY).setValue("");
  ACTIVE_SHEET.getRange(last_row+1, START_TIME).setValue("");
  ACTIVE_SHEET.getRange(last_row+1, END_TIME).setValue("");
  ACTIVE_SHEET.getRange(last_row+1, IS_COLENDER).setValue(false);
}

//↓HTMLでフォームを作成しようとしたがokta側でエラーが出ており数値を取得できない
/*function showTaskForm(){
  console.log("ShowWindow");
  const htmlOutput = HtmlService.createHtmlOutputFromFile(`taskForm`).setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `タスク入力`);
}

function processFormData(formData){
  console.log("a");
  ACTIVE_SHEET.getRange(5,5).setValue("梅野航大");
  if(formData){
    ACTIVE_SHEET.getRange(4,3).setValue(formData.taskName);
    ACTIVE_SHEET.getRange(4,4).setValue(formData.dueDate);
  }
}*/