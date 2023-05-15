///表ないの名前数え上げに関するファイル

const MAX_PEOPLE = 50;
const HEAD_ROW_OBOG = 14;
const HEAD_COLUMN_OBOG = 1;
const HEAD_ROW_ACTIVE = 14;
const HEAD_COLUMN_ACTIVE = 7;
const CELL_OBOG_NUM = "B12";
const HEAD_ROW_NON = 5;
const HEAD_COLUMN_NON = 2;
const CELL_ACTIVE_NUM = "H12";

///以下はbutton.gsから呼び出す形で
///name1, name2は無記入'', 名前1つ'須田', 名前2つ'須田・山﨑'の三種に対応して

//add_match(追加ボタン)に対応
//追加ボタン実行時にカウンタ増
function increase(name1, name2, type) {
  //名前を1つの配列に
  let name11 = name1.replace(" ", "").replace("　", "").split(/・|,|、/) //区切り文字："・,、"
  let name22 = name2.replace(" ", "").replace("　", "").split(/・|,|、/)
  let names = name11.concat(name22);
  names = names.filter(function(value){    
      return value !== ""      
      });  
  
  //OBOG、現役の人数取得
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

  //名前にカウンタ追加していく
  let n = 0;
  for(let name of names){
    if(name.includes("さん")){ //OBOGの場合
      var row = HEAD_ROW_OBOG;
      var column = HEAD_COLUMN_OBOG;
      var people = n_obog;
    }else{
      var row = HEAD_ROW_ACTIVE;
      var column = HEAD_COLUMN_ACTIVE;
      var people = n_active
    }

    let flag = 0;
    for(let i=0; i<people; i++){
      let get_name = list_sheet.getRange(row+i, column+1).getValue();
      if(name === get_name){ //一致するものがあった場合
        flag = 1;
        if(type === "エキシビ"){ //エキシビの場合
          var count = list_sheet.getRange(row+i, column+3).getValue();
          if(count===""){
            count = 1;
          }else{
            count = count+1;
          }
          list_sheet.getRange(row+i, column+3).setValue(count);
          break;
        }else{ //エキシビ以外の場合
          var count = list_sheet.getRange(row+i, column+2).getValue();
          if(count===""){
            count = 1;
          }else{
            count = count+1;
          }
          list_sheet.getRange(row+i, column+2).setValue(count);
          break;
        }
      }
    }
    //名前が一致するものがなかった場合
    if(flag==0){
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

//delete_match(削除ボタン)に対応
//削除ボタン実行時にカウンタ減
function decrease(name1, name2, type) {//名前1、名前2、試合種
  //名前を1つの配列に
  let name11 = name1.replace(" ", "").replace("　", "").split(/・|,|、/) //区切り文字："・,、"
  let name22 = name2.replace(" ", "").replace("　", "").split(/・|,|、/)
  let names = name11.concat(name22);
  names = names.filter(function(value){    
      return value !== ""      
      });  
  
  //OBOG、現役の人数取得
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

  //名前にカウンタ減らしていく
  let n = 0;
  for(let name of names){
    if(name.includes("さん")){ //OBOGの場合
      var row = HEAD_ROW_OBOG;
      var column = HEAD_COLUMN_OBOG;
      var people = n_obog;
    }else{
      var row = HEAD_ROW_ACTIVE;
      var column = HEAD_COLUMN_ACTIVE;
      var people = n_active
    }

    let flag = 0;
    for(let i=0; i<people; i++){
      let get_name = list_sheet.getRange(row+i, column+1).getValue();
      if(name === get_name){ //一致するものがあった場合
        flag = 1;
        if(type === "エキシビ"){ //エキシビの場合
          var count = list_sheet.getRange(row+i, column+3).getValue();
          if(count===""){
            count = 0;
          }else{
            count = count-1;
          }
          list_sheet.getRange(row+i, column+3).setValue(count);
          break;
        }else{ //エキシビ以外の場合
          var count = list_sheet.getRange(row+i, column+2).getValue();
          if(count===""){
            count = 0;
          }else{
            count = count-1;
          }
          list_sheet.getRange(row+i, column+2).setValue(count);
          break;
        }
      }
    }

    //名前が一致するものがなかった場合
    if(flag==0){
      if(type === "エキシビ"){ //エキシビの場合
        list_sheet.getRange(HEAD_ROW_NON+n,HEAD_COLUMN_NON,1,3).setValues([[[name],[''],['-1']]])
        n = n+1;
      }else{ //エキシビ以外の場合
        list_sheet.getRange(HEAD_ROW_NON+n,HEAD_COLUMN_NON,1,3).setValues([[[name],['-1'],['']]])
        n = n+1;
      }
      Browser.msgBox(name+"の名前はリストにありません。すぐに手動でカウント変更して削除してください");
    }
  }
}
