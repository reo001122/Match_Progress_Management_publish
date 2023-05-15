///ボタンによって実行される関数群のファイル

//ON,OFF切り替えボタン
function switch_onoff() {
  let range = sheet.getRange('C40');
  let value = range.getValue();
  if(value == 'OFF'){
    value = 'ON';
  }else{
    value = 'OFF';
  }
  range.setValue(value);
}

//コート数に対応してIDの振り分けボタン
function court_num() {
  if(judge_on_off()){
    range = sheet.getRange('C44');
    let n = range.getValue();
    if(n===''){
      n = MAX_COURTS;
    }
    n = Math.min(n, MAX_COURTS);
    for(let i=0; i<ROUNDS*MAX_COURTS; i++){
      let row = HEAD_ROW + (i%ROUNDS)*2 + 1;
      let column = HEAD_COLUMN + Math.floor(i/ROUNDS)*4;
      let range = sheet.getRange(row, column);

      if(i<ROUNDS*n){
        range.setValue(i);
        if(i%ROUNDS==0){
          sheet.getRange(5,column).setBackground(null)
        }
      }else{
        range.setValue('');
        if(i%ROUNDS==0){
          sheet.getRange(5,column).setBackground("#717375")
        }
      }
    }

    sheet.getRange('C45').setValue(n);
    range.setValue('');

  }else{
    Browser.msgBox('ON/OFFボタンでONにしてください')
  }
}

//試合追加するボタン
function add_match() {
  if(judge_on_off()){
    let getrange = sheet.getRange("F40:F43")
    let values = getrange.getValues();
    let id = values[0][0];
    let type = values[1][0];
    let player1 = values[2][0];
    let player2 = values[3][0];

    if(id!=='' && type!=='' && player1!=='' && player2!==''){
      let row =  HEAD_ROW + (id%ROUNDS)*2;
      let column = HEAD_COLUMN + Math.floor(id/ROUNDS)*4;
      let setrange = sheet.getRange(row, column+1, 2, 2);

      if(!setrange.isBlank()){ //移動先にすでに試合がある場合
        var result=Browser.msgBox("移動先にはすでに試合がありますが、移動して良いですか？",Browser.Buttons.OK_CANCEL);
        if(result=='cancel'){ //キャンセルする場合、以降実行せず
          return
        }
      }

      delete_match(id); //移動先の試合を削除
      setrange.setValues([[type, player1], ['', player2]]);
      getrange.setValues([[''],[''],[''],['']]);

      //カウンタ増加
      increase(player1, player2, type);
      // 選手状態を試合中に
      state_of_player(player1, player2, 2, 2, type);
    }else{
      Browser.msgBox('記入漏れがあります')
    }

  }else{
    Browser.msgBox('ON/OFFボタンでONにしてください')
  }
}

//試合を控えから表に移動するボタン
function move_reserve_match() {
  if(judge_on_off()){
    let getrange1 = sheet.getRange("J40:J41")
    let values1 = getrange1.getValues();
    let id_before = values1[0];
    let id_after = values1[1];

    if(id_before<=MAX_RESERVE_ID && (id_before!=''||id_before===0) && (id_after!=''||id_after===0)){
      let row_before = HEAD_ROW_RESERVE + id_before*2;
      let column_before = HEAD_COLUMN_RESERVE;
      let row_after = HEAD_ROW + (id_after%ROUNDS)*2;
      let column_after = HEAD_COLUMN + Math.floor(id_after/ROUNDS)*4;

      let getrange2 = sheet.getRange(row_before, column_before+1, 2, 2);
      let values2 = getrange2.getValues();
      
      let type = values2[0][0];
      let player1 = values2[0][1];
      let player2 = values2[1][1];
      values2[1][0] = '';
      //increase(player1, player2, type);
      
      let setrange = sheet.getRange(row_after, column_after+1, 2, 2);

      if(!setrange.isBlank()){ //移動先にすでに試合がある場合
        var result=Browser.msgBox("移動先にはすでに試合がありますが、移動して良いですか？",Browser.Buttons.OK_CANCEL);
        if(result=='cancel'){ //キャンセルする場合、以降実行せず
          return
        }
      }

      delete_match(id_after); //移動先の試合を削除
      setrange.setValues(values2); //移動先に入力
      getrange1.setValues([[''],['']]);
      getrange2.setValues([[[''],['']], [[''],['']]])
      //カウンタ増やす
      increase(player1, player2, type);
      // 選手状態を試合中に
      state_of_player(player1, player2, 2, 2, type);
      //控え表の調整
      reserve_up(id_before);
    }else{
      Browser.msgBox("IDが無効です")
    }
  }else{
    Browser.msgBox('ON/OFFボタンでONにしてください')
  }
}

//試合を表内で移動するボタン
function move_table_match() {
  if(judge_on_off()){
    let getrange1 = sheet.getRange("J45:J46");
    let values1 = getrange1.getValues();
    let id_before = values1[0];
    let id_after = values1[1];

    if((id_before!=''||id_before===0) && (id_after!=''||id_after===0)){
      let row_before = HEAD_ROW + (id_before%ROUNDS)*2;
      let column_before = HEAD_COLUMN + Math.floor(id_before/ROUNDS)*4;
      let row_after = HEAD_ROW + (id_after%ROUNDS)*2;
      let column_after = HEAD_COLUMN + Math.floor(id_after/ROUNDS)*4;

      let getrange2 = sheet.getRange(row_before, column_before+1, 2, 3);
      let values2 = getrange2.getValues();
      
      let setrange = sheet.getRange(row_after, column_after+1, 2, 3);

      if(!setrange.isBlank()){ //移動先にすでに試合がある場合
        var result=Browser.msgBox("移動先にはすでに試合がありますが、移動して良いですか？",Browser.Buttons.OK_CANCEL);
        if(result=='cancel'){ //キャンセルする場合、以降実行せず
          return
        }
      }
      delete_match(id_after); //まず移動先の試合を削除する
      setrange.setValues(values2); //移動先に入力
      getrange1.setValues([[''],['']]); 
      getrange2.setValues([[[''],[''],['']], [[''],[''],['']]])
      getrange2.setBackground(null);
    }else{
      Browser.msgBox("IDが無効です")
    }

  }else{
    Browser.msgBox('ON/OFFボタンでONにしてください')
  }
}

//試合を削除するボタン
function delete_match(id) {
  if(judge_on_off()){
    //削除するIDの取得
    if(id === undefined){//引数が空の場合(ボタンからの処理)
      let getrange = sheet.getRange("J50");
      var id = getrange.getValue();
      getrange.setValue('');
    }

    if(id===0 || id!=''){
      let row =  HEAD_ROW + (id%ROUNDS)*2;
      let column = HEAD_COLUMN + Math.floor(id/ROUNDS)*4;
      let setrange = sheet.getRange(row, column+1, 2, 3);
      let values = setrange.getValues();

      //すでにある試合を削除しようとすると警告を出す
      if(!setrange.isBlank()){
        var del_result = Browser.msgBox("本当にこの試合を削除しますか？",Browser.Buttons.OK_CANCEL);
        if(del_result == 'cancel'){
          return
        }
      }

      //カウンタ減らす
      let type = values[0][0];
      let player1 = values[0][1];
      let player2 = values[1][1];
      decrease(player1, player2, type);
      // 選手状態を待機に
      state_of_player(player1, player2, 0, 0, type);

      setrange.setValues([[[''],[''],['']], [[''],[''],['']]]);
      setrange.setBackground(null);
    }else{
      Browser.msgBox("IDが無効です")
    }

  }else{
    Browser.msgBox('ON/OFFボタンでONにしてください')
  }
}

//控えを解除するボタン
function cancel_reserve(id) {
  if(judge_on_off()){
    //削除するIDの取得
    if(id === undefined){//引数が空の場合(ボタンからの処理)
      let getrange = sheet.getRange("J60");
      var id = getrange.getValue();
      getrange.setValue('');
    }

    if((id===0 || id!='') && id <= MAX_RESERVE_ID){
      let row =  HEAD_ROW_RESERVE + 2*id;
      let column = HEAD_COLUMN_RESERVE;
      let setrange = sheet.getRange(row, column+1, 2, 2);
      let values = setrange.getValues();

      //すでにある試合を削除しようとすると警告を出す
      if(!setrange.isBlank()){
        var cancel_result = Browser.msgBox("本当にこの控えを削除しますか？",Browser.Buttons.OK_CANCEL);
        if(cancel_result == 'cancel'){
          return
        }
      }

      let type = values[0][0];
      let player1 = values[0][1];
      let player2 = values[1][1];
      // 選手状態を待機に
      state_of_player(player1, player2, 0, 0, type);

      setrange.setValues([[[''],['']], [[''],['']]]);
      // 控え表の調整
      reserve_up(id);
    }else {
      Browser.msgBox("IDが無効です");
    }
  }else {
    Browser.msgBox("ON/OFFボタンでONにしてください");
  }
}

//表上の全ての名前をカウント
function check_all() {
  if(judge_on_off()){
    let sheet_name = active_sheet_name.replace("名簿", "");
    let list_sheet_name = sheet_name + "名簿";
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
    let list_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(list_sheet_name);

    //値をクリアする
    list_sheet.getRange(HEAD_ROW_OBOG, HEAD_COLUMN_OBOG+2, MAX_PEOPLE, 2).clearContent();
    list_sheet.getRange(HEAD_ROW_ACTIVE, HEAD_COLUMN_ACTIVE+2, MAX_PEOPLE, 2).clearContent();

    for(let i=0; i<ROUNDS; i++){
      for(let j=0; j<MAX_COURTS; j++){
        let values = sheet.getRange(HEAD_ROW+2*i, HEAD_COLUMN+1+4*j, 2, 2).getValues();
        console.log(values);
        let type = values[0][0];
        let name1 = values[0][1];
        let name2 = values[1][1];
        
        if(name1 !== ''){
          increase(name1, name2, type);
        }
      }
    }

    Browser.msgBox("カウントが完了しました")
  }else{
    Browser.msgBox('ON/OFFボタンでONにしてください')
  }
}

function test(){
  let range = sheet.getRange(16, 4, 2, 3);
  console.log(range.isBlank());
}