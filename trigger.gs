///トリガーによって実行される関数群のファイル

//スコアが入力されたときに、時間入力する関数
//セル変更時トリガー
function input_time() {
  if(judge_on_off()){ 
    let sheet_name = sheet.getSheetName();
    if(sheet_name.includes("名簿") || sheet_name.includes("xyz")){ 
      return
    }
    //名簿系のシート・xyzをシート名に含む場合、以下全て実行しない

    let range = sheet.getRange('C45');
    let n = range.getValue();

    let changed_cell = sheet.getActiveCell();
    let changed_row = changed_cell.getRow();
    let changed_column = changed_cell.getColumn();
    changed_row = Math.floor(changed_row/2)*2;
    
    //スコア記入部分の更新時のみ以下を実行
    if(HEAD_ROW<=changed_row && changed_row<=HEAD_ROW+ROUNDS*2-1
    && HEAD_COLUMN<=changed_column && changed_column<=HEAD_COLUMN+n*4
    && changed_column%4==2){
      //まず終了していないか確認
      let score1 = sheet.getRange(changed_row, changed_column).getValue();
      let score2 = sheet.getRange(changed_row+1, changed_column).getValue();
      //終了した場合、finished入力
      if(check_isfinish(score1, score2)){
        let set_row = changed_row+1;
        let set_column = changed_column-2;
        let set_range = sheet.getRange(set_row, set_column);
        let set_range_bg = sheet.getRange(set_row-1, set_column, 2, 3);
        let values = set_range_bg.getValues();
        let type = values[0][0];
        let player1 = values[0][1];
        let player2 = values[1][1];

        set_range.setValue('finished');
        set_range.setFontColor("blue");
        set_range_bg.setBackground("#ebf6f7");
        // 選手状態を待機に戻す
        state_of_player(player1, player2, 0, 0, type);

      }else{//終了していない場合、時間入力
        let set_row = changed_row+1;
        let set_column = changed_column-2;
        let set_range = sheet.getRange(set_row, set_column);
        let set_range_bg = sheet.getRange(set_row-1, set_column, 2, 3);
        let values = set_range_bg.getValues();
        let type = values[0][0];
        let player1 = values[0][1];
        let player2 = values[1][1];

        let date = new Date();
        date = Utilities.formatDate(date, "Asia/Tokyo", "HH:mm:ss");
        set_range.setValue(date);
        set_range.setFontColor("black")
        set_range_bg.setBackground("#a0d8ef")
        //選手状態を試合中に戻す(もし間違えてfinishedにしてしまった時とか)
        state_of_player(player1, player2, 2, 2, type);
      }
    }
  }
}

//控えに入れた時刻を表示
//セル変更時トリガー
function input_time_reserve() {
  if(judge_on_off()){
    let sheet_name = sheet.getSheetName();
    if(sheet_name.includes("名簿") || sheet_name.includes("xyz")){ 
      return
    }
    //名簿系のシート・xyzをシート名に含む場合、以下全て実行しない

    let changed_cell = sheet.getActiveCell();
    let changed_row = changed_cell.getRow();
    let changed_column = changed_cell.getColumn();

    //控え記入時に入力
    if(HEAD_ROW_RESERVE <= changed_row 
    && changed_row<=HEAD_ROW_RESERVE+MAX_RESERVE_ID*2+1
    && HEAD_COLUMN_RESERVE+1<=changed_column
    && changed_column<=HEAD_COLUMN_RESERVE+2
    && !(changed_column==HEAD_COLUMN_RESERVE+1 && changed_row%2==1)){
      let set_row = Math.floor(changed_row/2)*2+1;
      let set_column = Math.floor(changed_column/2)*2;
      let set_range = sheet.getRange(set_row, set_column);
      let date = new Date();
      date = Utilities.formatDate(date, "Asia/Tokyo", "HH:mm:ss");
      set_range.setValue(date);
      // 以降、選手状態変更用
      // 選手状態を控えに
      let get_range = sheet.getRange(set_row-1, set_column, 2, 2);
      let values = get_range.getValues();
      let player1 = values[0][1];
      let player2 = values[1][1];
      let type = values[1][0];
      state_of_player(player1, player2, 1, 1, type);
    }else if(HEAD_ROW_REST <= changed_row 
    && changed_row <= HEAD_ROW_REST + MAX_REST
    && changed_column == HEAD_COLUMN_REST) {
      let set_row = changed_row;
      let set_column = changed_column+1;
      let set_range = sheet.getRange(set_row, set_column);
      let date = new Date();
      date = Utilities.formatDate(date, "Asia/Tokyo", "HH:mm:ss");
      set_range.setValue(date);
    }
  }
}


//1分ごとに時間超過チェックする関数
//時間主導のトリガー
function check_by_time() {
  if(judge_on_off()){
    let sheet_name = sheet.getSheetName();
    if(sheet_name.includes("名簿") || sheet_name.includes("xyz")){
      return
    }
    //名簿系のシート・xyzをシート名に含む場合、以下全て実行しない

    let range1 = sheet.getRange('C45');
    let n = range1.getValue();
    let now_time = new Date();

    for(let i=0; i<ROUNDS*n; i++){
      let row = HEAD_ROW + (i%ROUNDS)*2 + 1;
      let column = HEAD_COLUMN + Math.floor(i/ROUNDS)*4 + 1;
      range = sheet.getRange(row, column);
      var old_time = range.getValue();
     
      if(old_time != '' && old_time != 'finished'){
        //多分ここすごい怪しいことしてる。普通にやるとうまくいかなかった。
        old_time.setFullYear(now_time.getFullYear());
        old_time.setMonth(now_time.getMonth());
        old_time.setDate(now_time.getDate());
        old_time.setMinutes(old_time.getMinutes()+THRESHOLD_MIN);
        console.log([old_time, now_time]);
        if(old_time<now_time){ //時間が閾値超えてたら赤字に
          range.setFontColor("red");
        }
      }
    }
  }
}
