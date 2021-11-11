// chess pieces website : https://commons.wikimedia.org/wiki/Category:PNG_chess_pieces/Standard_transparent

var black_pawn = "https://upload.wikimedia.org/wikipedia/commons/c/cd/Chess_pdt60.png";
var white_pawn = "https://upload.wikimedia.org/wikipedia/commons/0/04/Chess_plt60.png";
var black_bishop = "https://upload.wikimedia.org/wikipedia/commons/8/81/Chess_bdt60.png";
var white_bishop = "https://upload.wikimedia.org/wikipedia/commons/9/9b/Chess_blt60.png";
var black_rook = "https://upload.wikimedia.org/wikipedia/commons/a/a0/Chess_rdt60.png";
var white_rook = "https://upload.wikimedia.org/wikipedia/commons/5/5c/Chess_rlt60.png";
var black_knight = "https://upload.wikimedia.org/wikipedia/commons/f/f1/Chess_ndt60.png";
var white_knight = "https://upload.wikimedia.org/wikipedia/commons/2/28/Chess_nlt60.png";
var black_queen = "https://upload.wikimedia.org/wikipedia/commons/a/af/Chess_qdt60.png";
var white_queen = "https://upload.wikimedia.org/wikipedia/commons/4/49/Chess_qlt60.png";
var black_king = "https://upload.wikimedia.org/wikipedia/commons/e/e3/Chess_kdt60.png";
var white_king = "https://upload.wikimedia.org/wikipedia/commons/3/3b/Chess_klt60.png";

var orig_backgrounds = {'F3':"#6aa84f", 'G3':"#d9ead3", 'H3':"#6aa84f", 'I3':"#d9ead3", 'J3':"#6aa84f", 'K3':"#d9ead3", 'L3':"#6aa84f", 'M3':"#d9ead3",
                  'F4':"#d9ead3", 'G4':"#6aa84f", 'H4':"#d9ead3", 'I4':"#6aa84f", 'J4':"#d9ead3", 'K4':"#6aa84f", 'L4':"#d9ead3", 'M4':"#6aa84f",
                  'F5':"#6aa84f", 'G5':"#d9ead3", 'H5':"#6aa84f", 'I5':"#d9ead3", 'J5':"#6aa84f", 'K5':"#d9ead3", 'L5':"#6aa84f", 'M5':"#d9ead3",
                  'F6':"#d9ead3", 'G6':"#6aa84f", 'H6':"#d9ead3", 'I6':"#6aa84f", 'J6':"#d9ead3", 'K6':"#6aa84f", 'L6':"#d9ead3", 'M6':"#6aa84f",
                  'F7':"#6aa84f", 'G7':"#d9ead3", 'H7':"#6aa84f", 'I7':"#d9ead3", 'J7':"#6aa84f", 'K7':"#d9ead3", 'L7':"#6aa84f", 'M7':"#d9ead3",
                  'F8':"#d9ead3", 'G8':"#6aa84f", 'H8':"#d9ead3", 'I8':"#6aa84f", 'J8':"#d9ead3", 'K8':"#6aa84f", 'L8':"#d9ead3", 'M8':"#6aa84f",
                  'F9':"#6aa84f", 'G9':"#d9ead3", 'H9':"#6aa84f", 'I9':"#d9ead3", 'J9':"#6aa84f", 'K9':"#d9ead3", 'L9':"#6aa84f", 'M9':"#d9ead3",
                  'F10':"#d9ead3", 'G10':"#6aa84f", 'H10':"#d9ead3", 'I10':"#6aa84f", 'J10':"#d9ead3", 'K10':"#6aa84f", 'L10':"#d9ead3", 'M10':"#6aa84f"};

function determine_next_position(piece, curr_loc){
  var boundary = ['F3', 'G3', 'H3', 'I3', 'J3', 'K3', 'L3', 'M3',
                  'F4', 'G4', 'H4', 'I4', 'J4', 'K4', 'L4', 'M4',
                  'F5', 'G5', 'H5', 'I5', 'J5', 'K5', 'L5', 'M5',
                  'F6', 'G6', 'H6', 'I6', 'J6', 'K6', 'L6', 'M6',
                  'F7', 'G7', 'H7', 'I7', 'J7', 'K7', 'L7', 'M7',
                  'F8', 'G8', 'H8', 'I8', 'J8', 'K8', 'L8', 'M8',
                  'F9', 'G9', 'H9', 'I9', 'J9', 'K9', 'L9', 'M9',
                  'F10', 'G10', 'H10', 'I10', 'J10', 'K10', 'L10', 'M10'];

  var alphabets = ["F", "G", "H", "I", "J", "K", "L", "M"];
  var numbers = ["3", "4", "5", "6", "7", "8", "9", "10"];
  var next_moves = {"up": "", "down": "", "left":"", "right":""};

  function splitter(cur_loc){
    let li = [];
    if(cur_loc.length == 2){
      for(let x in cur_loc){
        li.push(cur_loc[x]);
      }
    }
    else{
      li.push(cur_loc[0]);
      li.push(cur_loc[1]+cur_loc[2]);
    }
    return li;
  }
  var alpha_num = splitter(curr_loc);
  if(piece == 'https://upload.wikimedia.org/wikipedia/commons/c/cd/Chess_pdt60.png' ||
      piece == 'https://upload.wikimedia.org/wikipedia/commons/0/04/Chess_plt60.png'){
    
    // Evaluating Left and Right
    var cur_loc_alpha_index = alphabets.indexOf(alpha_num[0]);
    if(cur_loc_alpha_index != 0){
      var moveleft = alphabets[cur_loc_alpha_index-1]+alpha_num[1];
      Logger.log("PAWN go LEFT " + moveleft);
    }
    else{
      Logger.log("PAWN Cannot go LEFT");
      var moveleft = "";
    }
    if(cur_loc_alpha_index != alphabets.length-1){
      var moveright = alphabets[cur_loc_alpha_index+1]+alpha_num[1];
      Logger.log("PAWN go RIGHT " + moveright);
    }
    else{
      Logger.log("PAWN Cannot go RIGHT");
      var moveright = "";
    }

    // Evaluating Up and Down
    var cur_loc_alpha_index = numbers.indexOf(alpha_num[1]);
    if(cur_loc_alpha_index != 0){
      var moveup = alpha_num[0] + numbers[cur_loc_alpha_index-1];
      Logger.log("PAWN go UP " + moveup);
    }
    else{
      Logger.log("PAWN Cannot go UP");
      var moveup = "";
    }
    if(cur_loc_alpha_index != numbers.length-1){
      var movedown = alpha_num[0] + numbers[cur_loc_alpha_index+1];
      Logger.log("PAWN go DOWN " + movedown);
    }
    else{
      Logger.log("PAWN Cannot go DOWN");
      var movedown = "";
    }
    next_moves["left"] = moveleft;
    next_moves["right"] = moveright;
    next_moves["up"] = moveup;
    next_moves["down"] = movedown;
    next_moves['is_pawn'] = "true";
  }
  
  // For Rook
  if(piece == 'https://upload.wikimedia.org/wikipedia/commons/5/5c/Chess_rlt60.png' ||
      piece == 'https://upload.wikimedia.org/wikipedia/commons/a/a0/Chess_rdt60.png'){
        // Evaluating Left and Right
        var cur_loc_alpha_index = alphabets.indexOf(alpha_num[0]);
        if(cur_loc_alpha_index != 0){
          var moveleft = [];
          for(let i = cur_loc_alpha_index-1 ; i >= 0 ; i--){
            moveleft.push(alphabets[i]+alpha_num[1]);
          }
          Logger.log("ROOK go LEFT " + moveleft);
        }
        else{
          Logger.log("ROOK Cannot go LEFT");
          var moveleft = "";
        }
        if(cur_loc_alpha_index != alphabets.length-1){
          var moveright = [];
          for(let i = cur_loc_alpha_index+1 ; i <= alphabets.length-1 ; i++){
            moveright.push(alphabets[i]+alpha_num[1]);
          }
          Logger.log("ROOK go RIGHT " + moveright);
        }
        else{
          Logger.log("ROOK Cannot go RIGHT");
          var moveright = "";
        }

        // Evaluating Up and Down
        var cur_loc_alpha_index = numbers.indexOf(alpha_num[1]);
        if(cur_loc_alpha_index != 0){
          var moveup = [];
          for(let i = cur_loc_alpha_index-1 ; i >= 0 ; i--){
            moveup.push(alpha_num[0]+numbers[i]);
          }
          Logger.log("ROOK go UP " + moveup);
        }
        else{
          Logger.log("ROOK Cannot go UP");
          var moveup = "";
        }
        if(cur_loc_alpha_index != numbers.length-1){
          var movedown = [];
          for(let i = cur_loc_alpha_index+1 ; i <= alphabets.length-1 ; i++){
            movedown.push(alpha_num[0]+numbers[i]);
          }
          Logger.log("ROOK go DOWN " + movedown);
        }
        else{
          Logger.log("ROOK Cannot go DOWN");
          var movedown = "";
        }
      next_moves["left"] = moveleft;
      next_moves["right"] = moveright;
      next_moves["up"] = moveup;
      next_moves["down"] = movedown;
      next_moves['is_pawn'] = "false";
      }

  // For KNIGHT
  if(piece == "https://upload.wikimedia.org/wikipedia/commons/f/f1/Chess_ndt60.png" ||
      piece == "https://upload.wikimedia.org/wikipedia/commons/2/28/Chess_nlt60.png"){
        // Evaluating Left and Right
        var cur_loc_alpha_index = alphabets.indexOf(alpha_num[0]);
        var cur_loc_number_index = numbers.indexOf(alpha_num[1]);
        if(cur_loc_alpha_index != 0 && cur_loc_alpha_index != 1){ // main left condition
          var moveleft = [];
          if(cur_loc_number_index != 0 && cur_loc_number_index != 7){ //got two moves
            let two_lefts_alpha = alphabets[cur_loc_alpha_index - 2];
            moveleft.push(two_lefts_alpha + (parseInt(alpha_num[1]) - 1));
            moveleft.push(two_lefts_alpha + (parseInt(alpha_num[1]) + 1));
          }
          else if(cur_loc_number_index == 0){
            let two_lefts_alpha = alphabets[cur_loc_alpha_index - 2];
            moveleft.push(two_lefts_alpha + (parseInt(alpha_num[1]) + 1));
          }
          else if(cur_loc_number_index == 7){
            let two_lefts_alpha = alphabets[cur_loc_alpha_index - 2];
            moveleft.push(two_lefts_alpha + (parseInt(alpha_num[1]) - 1));
          }
          Logger.log("KNIGHT go LEFT " + moveleft);
        }
        else{
          Logger.log("KNIGHT Cannot go LEFT");
          var moveleft = "";
        }
        if(cur_loc_alpha_index != 7 && cur_loc_alpha_index != 6){ // main right condition
          var moveright = [];
          if(cur_loc_number_index != 0 && cur_loc_number_index != 7){ //got two moves
            let two_right_alpha = alphabets[cur_loc_alpha_index + 2];
            moveright.push(two_right_alpha + (parseInt(alpha_num[1]) - 1));
            moveright.push(two_right_alpha + (parseInt(alpha_num[1]) + 1));
          }
          else if(cur_loc_number_index == 0){
            let two_right_alpha = alphabets[cur_loc_alpha_index + 2];
            moveright.push(two_right_alpha + (parseInt(alpha_num[1]) + 1));
          }
          else if(cur_loc_number_index == 7){
            let two_right_alpha = alphabets[cur_loc_alpha_index + 2]; // e.g. H
            moveright.push(two_right_alpha + (parseInt(alpha_num[1]) - 1)); // e.g. H+10-1=H9
          }
          Logger.log("KNIGHT go RIGHT " + moveright);
        }
        else{
          Logger.log("KNIGHT Cannot go RIGHT");
          var moveright = "";
        }

        // Evaluating Up and Down
        if(cur_loc_number_index != 0 && cur_loc_number_index != 1){ // main right condition
          var moveup = [];
          if(cur_loc_alpha_index != 0 && cur_loc_alpha_index != 7){ //got two moves
            let two_ups_number = numbers[parseInt(cur_loc_number_index) - 2];
            moveup.push(alphabets[cur_loc_alpha_index+1]+two_ups_number);
            moveup.push(alphabets[cur_loc_alpha_index-1]+two_ups_number);
          }
          else if(cur_loc_alpha_index == 0){
            let two_ups_number = numbers[parseInt(cur_loc_number_index) - 2];
            moveup.push(alphabets[cur_loc_alpha_index+1]+two_ups_number);
          }
          else if(cur_loc_alpha_index == 7){
            let two_ups_number = numbers[parseInt(cur_loc_number_index) - 2];
            moveup.push(alphabets[cur_loc_alpha_index-1]+two_ups_number);
          }
          Logger.log("KNIGHT go UP " + moveup);
        }
        else{
          Logger.log("KNIGHT Cannot go UP");
          var moveup = "";
        }
        if(cur_loc_number_index != 7 && cur_loc_number_index != 6){ // main right condition
          var movedown = [];
          if(cur_loc_alpha_index != 0 && cur_loc_alpha_index != 7){ //got two moves
            let two_ups_number = numbers[parseInt(cur_loc_number_index) + 2];
            movedown.push(alphabets[cur_loc_alpha_index+1]+two_ups_number);
            movedown.push(alphabets[cur_loc_alpha_index-1]+two_ups_number);
          }
          else if(cur_loc_alpha_index == 0){
            let two_ups_number = numbers[parseInt(cur_loc_number_index) - 2];
            movedown.push(alphabets[cur_loc_alpha_index+1]+two_ups_number);
          }
          else if(cur_loc_alpha_index == 7){
            let two_ups_number = numbers[parseInt(cur_loc_number_index) - 2];
            movedown.push(alphabets[cur_loc_alpha_index-1]+two_ups_number);
          }
          Logger.log("KNIGHT go DOWN " + movedown);
        }
        else{
          Logger.log("KNIGHT Cannot go DOWN");
          var movedown = "";
        }
      next_moves["left"] = moveleft;
      next_moves["right"] = moveright;
      next_moves["up"] = moveup;
      next_moves["down"] = movedown;
      next_moves['is_pawn'] = "false";
      Logger.log("MOVES UNDER");
      Logger.log(next_moves);
    }
  return next_moves;
}

var board_positions = {'F3':black_rook, "G3":black_knight, "H3":black_bishop, "I3":black_queen, "J3":black_king, "K3":black_bishop, "L3":black_knight, "M3":black_rook,'F4':black_pawn, "G4":black_pawn, "H4":black_pawn, "I4":black_pawn, "J4":black_pawn, "K4":black_pawn, "L4":black_pawn, "M4":black_pawn,'F9':white_pawn, "G9":white_pawn, "H9":white_pawn, "I9":white_pawn, "J9":white_pawn, "K9":white_pawn, "L9":white_pawn, "M9":white_pawn,'F10':white_rook, "G10":white_knight, "H10":white_bishop, "I10":white_queen, "J10":white_king, "K10":white_bishop, "L10":white_knight, "M10":white_rook}

function start_board(){
  var initial_positions = {'F3':black_rook, "G3":black_knight, "H3":black_bishop, "I3":black_queen, "J3":black_king, "K3":black_bishop, "L3":black_knight, "M3":black_rook,'F4':black_pawn, "G4":black_pawn, "H4":black_pawn, "I4":black_pawn, "J4":black_pawn, "K4":black_pawn, "L4":black_pawn, "M4":black_pawn,'F9':white_pawn, "G9":white_pawn, "H9":white_pawn, "I9":white_pawn, "J9":white_pawn, "K9":white_pawn, "L9":white_pawn, "M9":white_pawn,'F10':white_rook, "G10":white_knight, "H10":white_bishop, "I10":white_queen, "J10":white_king, "K10":white_bishop, "L10":white_knight, "M10":white_rook}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formulaSheet = ss.getSheetByName("Chess");
  Object.keys(initial_positions).forEach(function(key) {
    var formulaCell = formulaSheet.getRange(key);
    formulaCell.setFormula('=IMAGE("' + initial_positions[key] + '")');
  });
}
var NEXT_MOVES = 'F997';
var IS_PAWN = "F999";
var PREV_POSITION = "F1000";

function test(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("others");
}


function pickup(){
  var moves = ['left', 'right', 'up', 'down'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chess");
  var eaoi = sheet.getActiveRange().getA1Notation();
  var formula = sheet.getActiveRange().getFormula();

  try{
    var link = formula.split('=IMAGE("')[1].split('")')[0];
    sheet.getRange(eaoi).setFormula('=IMAGE("")');
    sheet.getRange('B8').setFormula('=IMAGE("' + link + '")');
    var next_moves = determine_next_position(link, eaoi);
    Logger.log("MOVES");
    Logger.log(next_moves);
    // Coloring possbile moves
    if(next_moves['is_pawn'] == "true"){
      ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("pawn");
      var next_moves_setr = [];
      for(let i in moves){
        if(next_moves[moves[i]] != "" && next_moves[moves[i]]){
          next_moves_setr.push(next_moves[moves[i]]);
          sheet.getRange(next_moves[moves[i]]).setBackground("yellow");
        }
      }
      if(next_moves_setr.length != 0){
        ss.getSheetByName("Executions").getRange(NEXT_MOVES).setValue(next_moves_setr.toString());
        ss.getSheetByName("Executions").getRange(PREV_POSITION).setValue(eaoi);
      }
    }
    else{
      ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("others");
      var next_moves_setr = [];
      for(let i in moves){
        if(next_moves[moves[i]] != [] && next_moves[moves[i]]){
          for(let j in next_moves[moves[i]]){
            next_moves_setr.push(next_moves[moves[i]][j]);
            sheet.getRange(next_moves[moves[i]][j]).setBackground("yellow");
          }
        }
      }
      if(next_moves_setr.length != 0){
        ss.getSheetByName("Executions").getRange(NEXT_MOVES).setValue(next_moves_setr.toString());
        ss.getSheetByName("Executions").getRange(PREV_POSITION).setValue(eaoi);
      }
    }
    
  }// try end
  catch(err){
    Browser.msgBox("No Piece Found!" + err.message);
  }
}

function move(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chess");
  var aoi = sheet.getActiveRange().getA1Notation();

  // checking if moves is of PAWN
  var is_pawn = ss.getSheetByName("Executions").getRange(IS_PAWN).getValue();
  if(is_pawn == "pawn"){
    ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notpawn");
    var NEXT_MOVES_POS = ss.getSheetByName("Executions").getRange(NEXT_MOVES).getValue().split(",");
    if(NEXT_MOVES_POS != []){
      if(NEXT_MOVES_POS.includes(aoi)){
        var formula = sheet.getRange('B8').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange('B8').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
      }
      else{
        Browser.msgBox("Illegal Move!");
      }
    }
    else{
      Browser.msgBox("No Possbile Moves!");
    }
  }
  else if(is_pawn == "others"){
    ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notothers");
    var NEXT_MOVES_POS = ss.getSheetByName("Executions").getRange(NEXT_MOVES).getValue().split(",");
    if(NEXT_MOVES_POS != []){
      if(NEXT_MOVES_POS.includes(aoi)){
        var formula = sheet.getRange('B8').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange('B8').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
      }
      else{
        Browser.msgBox("Illegal Move!");
      }
    }
    else{
      Browser.msgBox("No Possbile Moves!");
    }
  }
}

function cancel(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eaoi = ss.getSheetByName("Executions").getRange(PREV_POSITION).getValue();
  var sheet = ss.getSheetByName("Chess");
  var formula = sheet.getRange('B8').getFormula();
  try{
    var link = formula.split('=IMAGE("')[1].split('")')[0];
    sheet.getRange('B8').setFormula('=IMAGE("")');
    sheet.getRange(eaoi).setFormula('=IMAGE("' + link + '")');
    var NEXT_MOVES_POS = ss.getSheetByName("Executions").getRange(NEXT_MOVES).getValue().split(",");
    for(let i in NEXT_MOVES_POS){
      sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
    }
  }
  catch{
    Browser.msgBox("Nothing to Revert.")
  }
}

function setEmails() {
  var emails = findItem();
  // SpreadsheetApp.getActiveSheet().getSheetByName("Chess").getRange('J10').setValue(emails.white);
  // SpreadsheetApp.getActiveSheet().getSheetByName("Chess").getRange('J11').setValue(emails.black);
  var ss = SpreadsheetApp.getActiveSheet();
  var dataSheet = ss.getSheetByName("Chess");
  dataSheet.getRange('J10').setValue(emails.white);
  dataSheet.getRange('J11').setValue(emails.black);
}

function setTimer(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Chess");
  var timerCell = dataSheet.getRange('J7');
  timerCell.setValue("00:00");
  timerCell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  timerCell.setFontSize(28);
}

function reset_timer(){
  setTimer();
}

function start_timer(){
  var d = new Date();
  var tick = d.getTime()
  var sourceSS = SpreadsheetApp.getActiveSpreadsheet();      //= Spreadsheet
  var dataSheet = sourceSS.getSheetByName("Executions");
  dataSheet.getRange('Z999').setValue(tick);
}

function stop_timer(){
  var d = new Date();
  var tock = d.getTime()
  var sourceSS = SpreadsheetApp.getActiveSpreadsheet();      //= Spreadsheet
  var dataSheet = sourceSS.getSheetByName("Executions");
  var tick = dataSheet.getRange('Z999').getValue();
  var minutes = Math.floor((tock-tick)/(24*3600));
  var seconds = Math.floor((tock-tick)/(24*60));
  var delta = minutes + " : " + seconds;
  Logger.log(delta);
  var timerCell = SpreadsheetApp.getActiveSheet().getRange('J7');
  timerCell.setValue(delta.toString());
  timerCell.setHorizontalAlignment("center").setVerticalAlignment("middle");
  timerCell.setFontSize(28);
}


function onOpen(e) {
  // Add a custom menu to the spreadsheet.
  // setEmails();
  // setTimer();
  start_board();
}

function sendMail(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();
  var sheetgId = ss.getActiveSheet().getSheetId();
  var sheetName = "Chess";

  var token = ScriptApp.getOAuthToken();

  var emails = findItem()

  var email = emails.white[0][0];
  var subject = "Opponent Player Made Move";
  var body = "PFA Current Board Positions Snap.";

  var url = "https://docs.google.com/spreadsheets/d/"+ssID+"/export?" + "format=xlsx" +  "&gid="+sheetgId+ "&portrait=true" + "&exportFormat=pdf";

  var result = UrlFetchApp.fetch(url, {
  headers: {
    'Authorization': 'Bearer ' +  token
  }
  });

  var contents = result.getContent();
  // Logger.log(emails.white[0][0]);
  MailApp.sendEmail(email,subject ,body, {name: 'Google Chess Engine', "cc": emails.black[0][0], attachments:[{fileName:sheetName+".pdf", content:contents, mimeType:"application//pdf"}]});
  Browser.msgBox("Opponent has been Notified.");
}

function findItem(){
  var sourceSS = SpreadsheetApp.getActiveSpreadsheet();      //= Spreadsheet
  // var id = sourceSS.getActiveSheet().getSheetId();
  // Logger.log(id);
  var dataSheet = sourceSS.getSheetByName("players");           //= Sheet
  var dataLastRow = dataSheet.getLastRow(); 
  Logger.log(dataLastRow);
  var white = dataSheet.getRange("B"+(dataLastRow)).getValues();
  var black = dataSheet.getRange("C"+(dataLastRow)).getValues();
  Logger.log({"white": white, "black": black});
  return {"white": white, "black": black}
}

function sheetnames() {
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) out.push(  sheets[i].getName()  )
  Logger.log(out);
}

function end_game() {
  var ui = SpreadsheetApp.getUi();
  
  // var black_button = CardService.newTextButton().setText("Black");
  // var white_button = CardService.newTextButton().setText("White");
  // var custom_buttons = CardService.newButtonSet().addButton(black_button).addButton(white_button);
  var result = ui.prompt("Are you sure, you want to end the game?");
  //Get the button that the user pressed.
  var button = result.getSelectedButton();
  
  if (button === ui.Button.OK) {
    Logger.log("The user clicked the [OK] button.");
    Logger.log(result.getResponseText());
    // var winner = result.getResponseText();
    // ui.alert("Congrats " + winner + "!");
  } else if (button === ui.Button.CLOSE) {
    Logger.log("The user clicked the [X] button and closed the prompt dialog."); 
    return 0
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();
  var file = DriveApp.getFileById(ssID);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();
  var sheetgId = ss.getActiveSheet().getSheetId();
  var sheetName = "Chess";

  var token = ScriptApp.getOAuthToken();

  var emails = findItem()

  var email = emails.white[0][0];
  var subject = "Game has Ended.";
  var body = "PFA Last Board Positions Snap.";

  var url = "https://docs.google.com/spreadsheets/d/"+ssID+"/export?" + "format=xlsx" +  "&gid="+sheetgId+ "&portrait=true" + "&exportFormat=pdf";

  var result = UrlFetchApp.fetch(url, {
  headers: {
    'Authorization': 'Bearer ' +  token
  }
  });

  var contents = result.getContent();
  // Logger.log(emails.white[0][0]);
  MailApp.sendEmail(email,subject ,body, {name: 'Google Chess Engine', "cc": emails.black[0][0], attachments:[{fileName:sheetName+".pdf", content:contents, mimeType:"application//pdf"}]});
  Browser.msgBox("The Game has Ended.");

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

// *****************************************************************************************
// AUTOMATING ENGINE CODES
// light red 1 light gray 2
var PAWN_H7 = 'H7';
function pawn_h7(){
  if(PAWN_H7 == 'H7'){
    SpreadsheetApp.getActiveSpreadsheet().getRange('H6').setBackground("yellow");
    SpreadsheetApp.getActiveSpreadsheet().getRange('H5').setBackground("yellow");
    // SpreadsheetApp.getActiveSpreadsheet().getRange('H5').assignScript("ClickMe");
  }
}

function SELECTED_RANGE() {
  // Logger.log(SpreadsheetApp.getActiveSpreadsheet().getActiveRange().getA1Notation());
  let pawn = "https://upload.wikimedia.org/wikipedia/commons/8/81/Chess_bdt60.png";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formulaSheet = ss.getSheetByName("Chess");
  var formulaCell = formulaSheet.getRange("D4");
  formulaCell.setFormula('=IMAGE("https://upload.wikimedia.org/wikipedia/commons/8/81/Chess_bdt60.png")');
  
}






