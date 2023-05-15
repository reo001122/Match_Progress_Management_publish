///mainファイル

const ROUNDS = 15; //最大ラウンド数
const MAX_COURTS = 8; //最大コート数
const HEAD_ROW = 6; //進行表の一番左上の行
const HEAD_COLUMN = 3; //進行表の一番左上の列
const HEAD_ROW_RESERVE = 54; //控え表の一番左上の行
const HEAD_COLUMN_RESERVE = 3; //控え表の一番左上の列
const MAX_RESERVE_ID = 15; //最大控え数
const THRESHOLD_MIN = 7; //時間超過の閾値
const HEAD_ROW_REST = 64; //レスト管理表の一番左上の行
const HEAD_COLUMN_REST = 8; //レスト管理表の一番左上の列
const MAX_REST = 5; //レストに入れる最大人数
const REST_TIME = 10; //レストに入れる時間
const ON_OFF = 'C40'; //on offボタンの位置


var sheet=SpreadsheetApp.getActiveSheet();
var active_sheet_name = sheet.getSheetName();

function judge_on_off() {
  let sheet_name = active_sheet_name.replace("名簿", "");
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  let range = sheet.getRange(ON_OFF);
  let on_off = range.getValue();
  return on_off == 'ON'
}

//試合終了かをbooleanで返す
function check_isfinish(score1, score2) {
  //score1,2の先頭の数字のみ切り出す（タイブレーク対応するため）
  score1 = Number(score1.toString().substring(0,1))
  score2 = Number(score2.toString().substring(0,1))
  if(score1==7 || score2==7){ //どちらかが7ゲーム目ゲット
    return true
  }else if(score1==6 && score1-score2>=2){ //score1が勝利
    return true
  }else if(score2==6 && score2-score1>=2){ //score2が勝利
    return true
  }else{
    return false
  }
}

//選手の状態を0:待機、1:控え、2:試合中の3つで色分け、名簿シートから見れるように
function state_of_player(name1, name2, state1, state2, type) {
//名前を1つの配列に
  let name11 = name1.replace(" ", "").replace("　", "").split(/・|,|、/) //区切り文字："・,、"
  let name22 = name2.replace(" ", "").replace("　", "").split(/・|,|、/)
  name11 = name11.filter(function(value){    
      return value !== ""      
      });
  name22 = name22.filter(function(value){
      return value !== ""
      });

  let list_sheet_name = active_sheet_name.replace("名簿", "") + "名簿";
  let list_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(list_sheet_name);
  let n_obog = list_sheet.getRange(CELL_OBOG_NUM).getValue();
  let n_active = list_sheet.getRange(CELL_ACTIVE_NUM).getValue();
  if(!isFinite(n_obog)){
    n_obog = MAX_PEOPLE;
  }
  if(!isFinite(n_active)){
    n_active = MAX_PEOPLE;
  }
  n_obog = Math.min(MAX_PEOPLE, n_obog);
  n_active = Math.min(MAX_PEOPLE, n_active);

  let n = 0;
  for(let name of name11){
    if(name.includes("さん")){ //OBOGの場合
      var row = HEAD_ROW_OBOG;
      var column = HEAD_COLUMN_OBOG;
      var people = n_obog;
    }else{
      var row = HEAD_ROW_ACTIVE;
      var column = HEAD_COLUMN_ACTIVE;
      var people = n_active
    }

    let flag1 = 0;
    for(let i=0; i<people; i++){
      let get_name = list_sheet.getRange(row+i, column+1).getValue();
      if(name === get_name){ //一致するものがあった場合
        flag1 = 1;
        let set_range = list_sheet.getRange(row+i, column, 1, 5);
        if (state1 === 0) { // 通常の待機状態
          set_range.setBackground('#ffffff');
        }else if (state1 === 1) {　// 控えに入っている状態
        　set_range.setBackground('#cee4ae');
        }else {　// 試合中の状態
          set_range.setBackground('#aacf53');
        }
      }
    }
    // 名前が一致するものがなかった場合
    if(flag1==0){
      if(type === "エキシビ"){ //エキシビの場合
        list_sheet.getRange(HEAD_ROW_NON+n,HEAD_COLUMN_NON,1,3).setValues([[[name],[''],['+1']]])
        n = n+1;
      }else{ //エキシビ以外の場合
        list_sheet.getRange(HEAD_ROW_NON+n,HEAD_COLUMN_NON,1,3).setValues([[[name],['+1'],['']]])
        n = n+1;
      }
      Browser.msgBox(name+"の名前はリストにありません。すぐに手動でカウント変更して削除してください");
    }
  }

  for(let name of name22){
    if(name.includes("さん")){ //OBOGの場合
      var row = HEAD_ROW_OBOG;
      var column = HEAD_COLUMN_OBOG;
      var people = n_obog;
    }else{
      var row = HEAD_ROW_ACTIVE;
      var column = HEAD_COLUMN_ACTIVE;
      var people = n_active
    }

    let flag2 = 0;
    for(let i=0; i<people; i++){
      let get_name = list_sheet.getRange(row+i, column+1).getValue();
      if(name === get_name){ //一致するものがあった場合
        flag2 = 1;
        let set_range = list_sheet.getRange(row+i, column, 1, 5);
        if (state2 === 0) { // 通常の待機状態
          set_range.setBackground('#ffffff');
        }else if (state2 === 1) {　// 控えに入っている状態
        　set_range.setBackground('#cee4ae');
        }else {　// 試合中の状態
          set_range.setBackground('#aacf53');
        }
      }
    }
    // 名前が一致するものがなかった場合
    if(flag2==0){
      if(type === "エキシビ"){ //エキシビの場合
        list_sheet.getRange(HEAD_ROW_NON+n,HEAD_COLUMN_NON,1,3).setValues([[[name],[''],['+1']]])
        n = n+1;
      }else{ //エキシビ以外の場合
        list_sheet.getRange(HEAD_ROW_NON+n,HEAD_COLUMN_NON,1,3).setValues([[[name],['+1'],['']]])
        n = n+1;
      }
      Browser.msgBox(name+"の名前はリストにありません。すぐに手動でカウント変更して削除してください");
    }
  }
}

//控え表から控えが消えた時に自動で残った控えを上に移動させるのに使う関数
function reserve_up(i) {
  if(judge_on_off()){
    for(let j = i; j < MAX_RESERVE_ID; j++) {
      // id i+1とid i+2の控えの内容を入れ替える
      let row_above = HEAD_ROW_RESERVE + 2*j;
      let column = HEAD_COLUMN_RESERVE
      let row_below = row_above + 2;
      let range_above = sheet.getRange(row_above, column+1, 2, 2);
      let values_above = range_above.getValues();
      let range_below = sheet.getRange(row_below, column+1, 2, 2);
      if(range_below.isBlank()) {
        break
      }
      let values_below = range_below.getValues();
      let tmp = values_above;
      values_above = values_below;
      values_below = tmp;
      range_above.setValues(values_above);
      range_below.setValues(values_below);
    }
  }else {
    Browser.msgBox('ON/OFFボタンでONにしてください');
  }
}