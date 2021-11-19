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

var references = {"https://upload.wikimedia.org/wikipedia/commons/c/cd/Chess_pdt60.png" : "black","https://upload.wikimedia.org/wikipedia/commons/8/81/Chess_bdt60.png" : "black","https://upload.wikimedia.org/wikipedia/commons/a/a0/Chess_rdt60.png" : "black","https://upload.wikimedia.org/wikipedia/commons/f/f1/Chess_ndt60.png" : "black","https://upload.wikimedia.org/wikipedia/commons/a/af/Chess_qdt60.png" : "black","https://upload.wikimedia.org/wikipedia/commons/e/e3/Chess_kdt60.png" : "black", "https://upload.wikimedia.org/wikipedia/commons/0/04/Chess_plt60.png" : "white","https://upload.wikimedia.org/wikipedia/commons/9/9b/Chess_blt60.png" : "white","https://upload.wikimedia.org/wikipedia/commons/5/5c/Chess_rlt60.png" : "white","https://upload.wikimedia.org/wikipedia/commons/2/28/Chess_nlt60.png" : "white","https://upload.wikimedia.org/wikipedia/commons/4/49/Chess_qlt60.png" : "white","https://upload.wikimedia.org/wikipedia/commons/3/3b/Chess_klt60.png" : "white"}

var orig_backgrounds = {'F3':"#6aa84f", 'G3':"#d9ead3", 'H3':"#6aa84f", 'I3':"#d9ead3", 'J3':"#6aa84f", 'K3':"#d9ead3", 'L3':"#6aa84f", 'M3':"#d9ead3",
                  'F4':"#d9ead3", 'G4':"#6aa84f", 'H4':"#d9ead3", 'I4':"#6aa84f", 'J4':"#d9ead3", 'K4':"#6aa84f", 'L4':"#d9ead3", 'M4':"#6aa84f",
                  'F5':"#6aa84f", 'G5':"#d9ead3", 'H5':"#6aa84f", 'I5':"#d9ead3", 'J5':"#6aa84f", 'K5':"#d9ead3", 'L5':"#6aa84f", 'M5':"#d9ead3",
                  'F6':"#d9ead3", 'G6':"#6aa84f", 'H6':"#d9ead3", 'I6':"#6aa84f", 'J6':"#d9ead3", 'K6':"#6aa84f", 'L6':"#d9ead3", 'M6':"#6aa84f",
                  'F7':"#6aa84f", 'G7':"#d9ead3", 'H7':"#6aa84f", 'I7':"#d9ead3", 'J7':"#6aa84f", 'K7':"#d9ead3", 'L7':"#6aa84f", 'M7':"#d9ead3",
                  'F8':"#d9ead3", 'G8':"#6aa84f", 'H8':"#d9ead3", 'I8':"#6aa84f", 'J8':"#d9ead3", 'K8':"#6aa84f", 'L8':"#d9ead3", 'M8':"#6aa84f",
                  'F9':"#6aa84f", 'G9':"#d9ead3", 'H9':"#6aa84f", 'I9':"#d9ead3", 'J9':"#6aa84f", 'K9':"#d9ead3", 'L9':"#6aa84f", 'M9':"#d9ead3",
                  'F10':"#d9ead3", 'G10':"#6aa84f", 'H10':"#d9ead3", 'I10':"#6aa84f", 'J10':"#d9ead3", 'K10':"#6aa84f", 'L10':"#d9ead3", 'M10':"#6aa84f"};

var black_occupied_positions = ['R3', 'S3', 'T3', 'U3',
                                'R4', 'S4', 'T4', 'U4',
                                'R5', 'S5', 'T5', 'U5',
                                'R6', 'S6', 'T6', 'U6'];
var white_occupied_positions = ['R7', 'S7', 'T7', 'U7',
                                'R8', 'S8', 'T8', 'U8',
                                'R9', 'S9', 'T9', 'U9',
                                'R10', 'S10', 'T10', 'U10'];


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
    var cur_loc_number_index = numbers.indexOf(alpha_num[1]);
    var pawn_red_moves = [];
    if(cur_loc_alpha_index != 0 && cur_loc_number_index != 0){ // main left up 
      pawn_red_moves.push(alphabets[cur_loc_alpha_index-1] + numbers[cur_loc_number_index-1]);
      
    }
    if(cur_loc_alpha_index != 0){
      var moveleft = alphabets[cur_loc_alpha_index-1]+alpha_num[1];
      Logger.log("PAWN go LEFT " + moveleft);
    }
    else{
      Logger.log("PAWN Cannot go LEFT");
      var moveleft = "";
    }

    if(cur_loc_alpha_index != 7 && cur_loc_number_index != 0){ // main right up condition
      pawn_red_moves.push(alphabets[cur_loc_alpha_index+1] + numbers[cur_loc_number_index-1]);
      
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
    if(cur_loc_alpha_index != 0 && cur_loc_number_index != 7){ // main down left condition
      pawn_red_moves.push(alphabets[cur_loc_alpha_index-1] + numbers[cur_loc_number_index+1]);
      
    }
    if(cur_loc_number_index != 0){
      var moveup = alpha_num[0] + numbers[cur_loc_number_index-1];
      Logger.log("PAWN go UP " + moveup);
    }
    else{
      Logger.log("PAWN Cannot go UP");
      var moveup = "";
    }
    if(cur_loc_alpha_index != 7 && cur_loc_number_index != 7){ // main down right condition
      pawn_red_moves.push(alphabets[cur_loc_alpha_index+1] + numbers[cur_loc_number_index+1]);
      
    }
    if(cur_loc_number_index != 7){
      var movedown = alpha_num[0] + numbers[cur_loc_number_index+1];
      Logger.log("PAWN go DOWN " + movedown);
    }
    else{
      Logger.log("PAWN Cannot go DOWN");
      var movedown = "";
    }
    if(pawn_red_moves == []){
      Logger.log("No Red Moves for PAWN");
    }
    next_moves["left"] = "";
    next_moves["right"] = "";
    next_moves["up"] = moveup;
    next_moves["down"] = movedown;
    next_moves["reds"] = pawn_red_moves;
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
    }

    // For BISHOP
  if(piece == "https://upload.wikimedia.org/wikipedia/commons/8/81/Chess_bdt60.png" ||
      piece == "https://upload.wikimedia.org/wikipedia/commons/9/9b/Chess_blt60.png"){
        // Evaluating Left and Right UPs
        var cur_loc_alpha_index = alphabets.indexOf(alpha_num[0]);
        var cur_loc_number_index = numbers.indexOf(alpha_num[1]);
        if(cur_loc_alpha_index != 0 && cur_loc_number_index != 0){ // main left up condition H.index != 0
          var moveleft = [];
          var j = 1;
          for(let i = cur_loc_alpha_index-1; i >= 0 ; i--){ // i = 2-1 = 1
            if(alphabets[i] != undefined && numbers[cur_loc_number_index-j] != undefined){
              moveleft.push(alphabets[i]+numbers[cur_loc_number_index-j])
              j = j + 1;
            }
          }
          Logger.log("BISHOP go LEFT " + moveleft);
        }
        else{
          Logger.log("BISHOP Cannot go LEFT");
          var moveleft = "";
        }
        if(cur_loc_alpha_index != 7 && cur_loc_number_index != 0){ // main right up condition
          var moveright = [];
          var j = 1;
          for(let i = cur_loc_alpha_index+1; i <= 7 ; i++){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index-j] != undefined){
              moveright.push(alphabets[i]+numbers[cur_loc_number_index-j]);
              j = j + 1;
            }
          }
          Logger.log("BISHOP go RIGHT " + moveright);
        }
        else{
          Logger.log("BISHOP Cannot go RIGHT");
          var moveright = "";
        }

        // Evaluating Left and Right DOWNs
        if(cur_loc_alpha_index != 0 && cur_loc_number_index != 7){ // main left up condition
          var moveup = [];
          var j = 1;
          for(let i = cur_loc_alpha_index-1; i >= 0 ; i--){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index+j] != undefined){
              moveup.push(alphabets[i]+numbers[cur_loc_number_index+j]);
              j = j + 1;
            }
          }
          Logger.log("BISHOP go UP " + moveup);
        }
        else{
          Logger.log("BISHOP Cannot go UP");
          var moveup = "";
        }
        if(cur_loc_alpha_index != 7 && cur_loc_number_index != 7){ // main left up condition
          var movedown = [];
          var j = 1;
          for(let i = cur_loc_alpha_index+1; i <= 7 ; i++){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index+j] != undefined){
              movedown.push(alphabets[i]+numbers[cur_loc_number_index+j])
              j = j + 1;
            }
          }
          Logger.log("BISHOP go DOWN " + movedown);
        }
        else{
          Logger.log("BISHOP Cannot go DOWN");
          var movedown = "";
        }
      next_moves["left"] = moveleft;
      next_moves["right"] = moveright;
      next_moves["up"] = moveup;
      next_moves["down"] = movedown;
      next_moves['is_pawn'] = "false";
    }

    // For QUEEN (COMBO OF BISHOP AND ROOK)
  if(piece == "https://upload.wikimedia.org/wikipedia/commons/a/af/Chess_qdt60.png" ||
      piece == "https://upload.wikimedia.org/wikipedia/commons/4/49/Chess_qlt60.png"){
        // Evaluating Left and Right UPs
        var cur_loc_alpha_index = alphabets.indexOf(alpha_num[0]);
        var cur_loc_number_index = numbers.indexOf(alpha_num[1]);
        var moveleft = [];
        var moveright = [];
        var movedown = [];
        var moveup = [];
        var moveleft2 = [];
        var moveright2 = [];
        var movedown2 = [];
        var moveup2 = [];
        if(cur_loc_alpha_index != 0 && cur_loc_number_index != 0){ // main left up 
          var j = 1;
          for(let i = cur_loc_alpha_index-1; i >= 0 ; i--){ // i = 2-1 = 1
            if(alphabets[i] != undefined && numbers[cur_loc_number_index-j] != undefined){
              moveleft.push(alphabets[i]+numbers[cur_loc_number_index-j])
              j = j + 1;
            }
          }
          Logger.log("QUEEN go UP LEFT " + moveleft);
        }
        if(cur_loc_alpha_index != 0){
          for(let i = cur_loc_alpha_index-1 ; i >= 0 ; i--){
            moveleft2.push(alphabets[i]+alpha_num[1]);
          }
          Logger.log("QUEEN go LEFT " + moveleft2);
        }
        if(moveleft == []){
          Logger.log("QUEEN Cannot go LEFT");
          var moveleft = "";
        }
        if(moveleft2 == []){
          Logger.log("QUEEN Cannot go LEFT 2");
          var moveleft2 = "";
        }
        if(cur_loc_alpha_index != 7 && cur_loc_number_index != 0){ // main right up condition
          var j = 1;
          for(let i = cur_loc_alpha_index+1; i <= 7 ; i++){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index-j] != undefined){
              moveright.push(alphabets[i]+numbers[cur_loc_number_index-j]);
              j = j + 1;
            }
          }
          Logger.log("QUEEN go UP RIGHT " + moveright);
        }
        if(cur_loc_alpha_index != alphabets.length-1){
          for(let i = cur_loc_alpha_index+1 ; i <= alphabets.length-1 ; i++){
            moveright2.push(alphabets[i]+alpha_num[1]);
          }
          Logger.log("QUEEN go RIGHT " + moveright2);
        }
        if(moveright == []){
          Logger.log("QUEEN Cannot go RIGHT");
          var moveright = "";
        }
        if(moveright2 == []){
          Logger.log("QUEEN Cannot go RIGHT 2");
          var moveright2 = "";
        }

        // Evaluating Left and Right DOWNs
        if(cur_loc_alpha_index != 0 && cur_loc_number_index != 7){ // main left up condition
          var j = 1;
          for(let i = cur_loc_alpha_index-1; i >= 0 ; i--){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index+j] != undefined){
              moveup.push(alphabets[i]+numbers[cur_loc_number_index+j]);
              j = j + 1;
            }
          }
          Logger.log("QUEEN go UP left" + moveup);
        }
        if(cur_loc_number_index != 0){
          for(let i = cur_loc_number_index-1 ; i >= 0 ; i--){
            moveup2.push(alpha_num[0]+numbers[i]);
          }
          Logger.log("QUEEN go UP " + moveup2);
        }
        if(moveup == []){
          Logger.log("QUEEN Cannot go UP");
          var moveup = "";
        }
        if(moveup2 == []){
          Logger.log("QUEEN Cannot go UP 2");
          var moveup2 = "";
        }
        if(cur_loc_alpha_index != 7 && cur_loc_number_index != 7){ // main left up condition
          var j = 1;
          for(let i = cur_loc_alpha_index+1; i <= 7 ; i++){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index+j] != undefined){
              movedown.push(alphabets[i]+numbers[cur_loc_number_index+j])
              j = j + 1;
            }
          }
          Logger.log("QUEEN go DOWN right" + movedown);
        }
        if(cur_loc_number_index != numbers.length-1){
          for(let i = cur_loc_number_index+1 ; i <= alphabets.length-1 ; i++){
            movedown2.push(alpha_num[0]+numbers[i]);
          }
          Logger.log("QUEEN go DOWN " + movedown2);
        }
        if(movedown == []){
          Logger.log("QUEEN Cannot go DOWN");
          var movedown = "";
        }
        if(movedown2 == []){
          Logger.log("QUEEN Cannot go DOWN 2");
          var movedown2 = "";
        }
      next_moves["left"] = moveleft;
      next_moves["right"] = moveright;
      next_moves["up"] = moveup;
      next_moves["down"] = movedown;
      next_moves["left2"] = moveleft2;
      next_moves["right2"] = moveright2;
      next_moves["up2"] = moveup2;
      next_moves["down2"] = movedown2;
      next_moves['is_pawn'] = "queen";
    }

    // For KING (COMBO OF BISHOP AND ROOK till ONE STEP)
  if(piece == "https://upload.wikimedia.org/wikipedia/commons/e/e3/Chess_kdt60.png" ||
      piece == "https://upload.wikimedia.org/wikipedia/commons/3/3b/Chess_klt60.png"){
        // Evaluating Left and Right UPs
        var cur_loc_alpha_index = alphabets.indexOf(alpha_num[0]);
        var cur_loc_number_index = numbers.indexOf(alpha_num[1]);
        var moveleft = [];
        var moveright = [];
        var movedown = [];
        var moveup = [];
        if(cur_loc_alpha_index != 0 && cur_loc_number_index != 0){ // main left up 
          var j = 1;
          for(let i = cur_loc_alpha_index-1; i >= cur_loc_alpha_index-1 ; i--){ // i = 2-1 = 1
            if(alphabets[i] != undefined && numbers[cur_loc_number_index-j] != undefined){
              moveleft.push(alphabets[i]+numbers[cur_loc_number_index-j])
              j = j + 1;
            }
          }
          Logger.log("KING go UP LEFT " + moveleft);
        }
        if(cur_loc_alpha_index != 0){
          for(let i = cur_loc_alpha_index-1 ; i >= cur_loc_alpha_index-1 ; i--){
            moveleft.push(alphabets[i]+alpha_num[1]);
          }
          Logger.log("KING go LEFT " + moveleft);
        }
        if(moveleft == []){
          Logger.log("KING Cannot go LEFT");
          var moveleft = "";
        }
        if(cur_loc_alpha_index != 7 && cur_loc_number_index != 0){ // main right up condition
          var j = 1;
          for(let i = cur_loc_alpha_index+1; i <= cur_loc_alpha_index+1 ; i++){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index-j] != undefined){
              moveright.push(alphabets[i]+numbers[cur_loc_number_index-j]);
              j = j + 1;
            }
          }
          Logger.log("KING go UP RIGHT " + moveright);
        }
        if(cur_loc_alpha_index != alphabets.length-1){
          for(let i = cur_loc_alpha_index+1 ; i <= alphabets.length-cur_loc_alpha_index+1 ; i++){
            moveright.push(alphabets[i]+alpha_num[1]);
          }
          Logger.log("KING go RIGHT " + moveright);
        }
        if(moveright == []){
          Logger.log("KING Cannot go RIGHT");
          var moveright = "";
        }

        // Evaluating Left and Right DOWNs
        if(cur_loc_alpha_index != 0 && cur_loc_number_index != 7){ // main left up condition
          var j = 1;
          for(let i = cur_loc_alpha_index-1; i >= cur_loc_alpha_index-1 ; i--){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index+j] != undefined){
              moveup.push(alphabets[i]+numbers[cur_loc_number_index+j]);
              j = j + 1;
            }
          }
          Logger.log("KING go UP left" + moveup);
        }
        if(cur_loc_number_index != 0){
          for(let i = cur_loc_number_index-1 ; i >= cur_loc_number_index-1 ; i--){
            moveup.push(alpha_num[0]+numbers[i]);
          }
          Logger.log("KING go UP " + moveup);
        }
        if(moveup == []){
          Logger.log("KING Cannot go UP");
          var moveup = "";
        }
        if(cur_loc_alpha_index != 7 && cur_loc_number_index != 7){ // main left up condition
          var j = 1;
          for(let i = cur_loc_alpha_index+1; i <= cur_loc_alpha_index+1 ; i++){
            if(alphabets[i] != undefined && numbers[cur_loc_number_index+j] != undefined){
              movedown.push(alphabets[i]+numbers[cur_loc_number_index+j])
              j = j + 1;
            }
          }
          Logger.log("KING go DOWN right" + movedown);
        }
        if(cur_loc_number_index != numbers.length-1){
          for(let i = cur_loc_number_index+1 ; i <= cur_loc_number_index+1 ; i++){
            movedown.push(alpha_num[0]+numbers[i]);
          }
          Logger.log("KING go DOWN " + movedown);
        }
        if(movedown == []){
          Logger.log("KING Cannot go DOWN");
          var movedown = "";
        }
      next_moves["left"] = moveleft;
      next_moves["right"] = moveright;
      next_moves["up"] = moveup;
      next_moves["down"] = movedown;
      next_moves['is_pawn'] = "false";
    }

  return next_moves;
}

var board_positions = {'F3':black_rook, "G3":black_knight, "H3":black_bishop, "I3":black_queen, "J3":black_king, "K3":black_bishop, "L3":black_knight, "M3":black_rook,'F4':black_pawn, "G4":black_pawn, "H4":black_pawn, "I4":black_pawn, "J4":black_pawn, "K4":black_pawn, "L4":black_pawn, "M4":black_pawn,'F9':white_pawn, "G9":white_pawn, "H9":white_pawn, "I9":white_pawn, "J9":white_pawn, "K9":white_pawn, "L9":white_pawn, "M9":white_pawn,'F10':white_rook, "G10":white_knight, "H10":white_bishop, "I10":white_queen, "J10":white_king, "K10":white_bishop, "L10":white_knight, "M10":white_rook}

var NEXT_MOVES = 'F997';
var RED_MOVES = 'F996';
var IS_PAWN = "F999";
var PREV_POSITION = "F1000";
var BLACK_OCCUPIED_POS = "F995";
var WHITE_OCCUPIED_POS = "F994";
var IS_GAME_STARTED = "F993";

function start_board(){
  var initial_positions = {'F3':black_rook, "G3":black_knight, "H3":black_bishop, "I3":black_queen, "J3":black_king, "K3":black_bishop, "L3":black_knight, "M3":black_rook,'F4':black_pawn, "G4":black_pawn, "H4":black_pawn, "I4":black_pawn, "J4":black_pawn, "K4":black_pawn, "L4":black_pawn, "M4":black_pawn,'F9':white_pawn, "G9":white_pawn, "H9":white_pawn, "I9":white_pawn, "J9":white_pawn, "K9":white_pawn, "L9":white_pawn, "M9":white_pawn,'F10':white_rook, "G10":white_knight, "H10":white_bishop, "I10":white_queen, "J10":white_king, "K10":white_bishop, "L10":white_knight, "M10":white_rook}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formulaSheet = ss.getSheetByName("Chess");
  Object.keys(initial_positions).forEach(function(key) {
    var formulaCell = formulaSheet.getRange(key);
    formulaCell.setFormula('=IMAGE("' + initial_positions[key] + '")');
  });

  var score_postions = {"B4": white_pawn, "C4" : black_pawn, "B5": white_knight, "C5" : black_knight, "B6": white_bishop, "C6" : black_bishop, "B7": white_rook, "C7" : black_rook, "B8": white_queen, "C8" : black_queen, "B9": white_king, "C9" : black_king}
  Object.keys(score_postions).forEach(function(key) {
    var formulaCell = formulaSheet.getRange(key);
    formulaCell.setFormula('=IMAGE("' + score_postions[key] + '")');
  });
  formulaSheet.getRange("P3").setFormula('=IMAGE("")');
  for(let wp in white_occupied_positions){
    formulaSheet.getRange(white_occupied_positions[wp]).setFormula('=IMAGE("")');
  }
  for(let bp in black_occupied_positions){
    formulaSheet.getRange(black_occupied_positions[bp]).setFormula('=IMAGE("")');
  }
  ss.getSheetByName("Executions").getRange(WHITE_OCCUPIED_POS).setValue(white_occupied_positions[0]);
  ss.getSheetByName("Executions").getRange(BLACK_OCCUPIED_POS).setValue(black_occupied_positions[0]);
  ss.getSheetByName("Executions").getRange(IS_GAME_STARTED).setValue("yes");
  var empty_board_positions = ['F5', 'F6', 'F7', 'F8',
                                'G5', 'G6', 'G7', 'G8',
                                'H5', 'H6', 'H7', 'H8',
                                'I5', 'I6', 'I7', 'I8',
                                'J5', 'J6', 'J7', 'J8',
                                'K5', 'K6', 'K7', 'K8',
                                'L5', 'L6', 'L7', 'L8',
                                'M5', 'M6', 'M7', 'M8'];
  for(let ebp in empty_board_positions){
    ss.getSheetByName("Chess").getRange(empty_board_positions[ebp]).setFormula('=IMAGE("")');
  }
  ss.getSheetByName("Chess").getRange('A16').setValue('');
  ss.getSheetByName("Chess").getRange('D16').setValue('');
}

function get_occupying_position(piece){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wop = ss.getSheetByName("Executions").getRange(WHITE_OCCUPIED_POS).getValue();
  var bop = ss.getSheetByName("Executions").getRange(BLACK_OCCUPIED_POS).getValue();
  if(piece == "black"){
    ss.getSheetByName("Executions").getRange(WHITE_OCCUPIED_POS).setValue(white_occupied_positions[white_occupied_positions.indexOf(wop) + 1]);
    return wop;
  }
  else{
    ss.getSheetByName("Executions").getRange(BLACK_OCCUPIED_POS).setValue(black_occupied_positions[black_occupied_positions.indexOf(bop) + 1]);
    return bop;
  }
}

function test(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chess");
  var temp_formula = sheet.getRange('I10').getFormula();
  temp_formula = temp_formula.split('=IMAGE("')[1].split('")')[0];
  Logger.log(temp_formula);
  if (!temp_formula.endsWith('.png')){
    Logger.log("YES");
  }
  else{
    Logger.log("NO");
  }
}


function pickup(){
  var moves = ['left', 'right', 'up', 'down'];
  var queen_moves = ['left', 'right', 'up', 'down', 'left2', 'right2', 'up2', 'down2'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chess");
  var eaoi = sheet.getActiveRange().getA1Notation();
  var formula = sheet.getActiveRange().getFormula();

  try{
    var link = formula.split('=IMAGE("')[1].split('")')[0];
    sheet.getRange(eaoi).setFormula('=IMAGE("")');
    sheet.getRange('P3').setFormula('=IMAGE("' + link + '")');
    var possible_moves = [];
    var red_moves = [];
    var next_moves = determine_next_position(link, eaoi);
    if(next_moves['is_pawn'] == "queen"){
      for(let i in queen_moves){
        if(next_moves[queen_moves[i]] != "" && next_moves[queen_moves[i]]){
          if(next_moves[queen_moves[i]] instanceof Array){
            for(let j in next_moves[queen_moves[i]]){
              try{
                var temp_formula = sheet.getRange(next_moves[queen_moves[i]][j]).getFormula();
                temp_formula = temp_formula.split('=IMAGE("')[1].split('")')[0];
                if (!temp_formula.endsWith('.png')){
                  possible_moves.push(next_moves[queen_moves[i]][j]);  
                }
                else{
                  if((references[link] == "white" && references[temp_formula] == "black")||
                  (references[link] == "black" && references[temp_formula] == "white")){
                    red_moves.push(next_moves[queen_moves[i]][j]);
                  }
                  break;
                }
              }
              catch{
                possible_moves.push(next_moves[queen_moves[i]][j]);
                if((references[link] == "white" && references[temp_formula] == "black")||
                  (references[link] == "black" && references[temp_formula] == "white")){
                    red_moves.push(next_moves[queen_moves[i]][j]);
                }
              }
            }
          }
        }
      }
    }
    else{
      for(let i in moves){
        if(next_moves[moves[i]] != "" && next_moves[moves[i]]){
          if(next_moves[moves[i]] instanceof Array){
            for(let j in next_moves[moves[i]]){
              try{
                var temp_formula = sheet.getRange(next_moves[moves[i]][j]).getFormula();
                temp_formula = temp_formula.split('=IMAGE("')[1].split('")')[0];
                if (!temp_formula.endsWith('.png')){
                  possible_moves.push(next_moves[moves[i]][j]);  
                }
                else{
                  if((references[link] == "white" && references[temp_formula] == "black")||
                  (references[link] == "black" && references[temp_formula] == "white")){
                    red_moves.push(next_moves[queen_moves[i]][j]);
                  }
                  break;
                }
              }
              catch{
                if((references[link] == "white" && references[temp_formula] == "black")||
                  (references[link] == "black" && references[temp_formula] == "white")){
                    red_moves.push(next_moves[queen_moves[i]][j]);
                }
                possible_moves.push(next_moves[moves[i]][j]);
              }
            }
          }
          else{
            //
          }
        }
      }
      if(next_moves['is_pawn'] == "true"){
        if(references[link] == "white"){
          try{
            var tem_formula = sheet.getRange(next_moves["up"]).getFormula();
            tem_formula = tem_formula.split('=IMAGE("')[1].split('")')[0];
            if (!tem_formula.endsWith('.png')){
              possible_moves.push(next_moves["up"]);
              if(ss.getSheetByName("Executions").getRange(IS_GAME_STARTED).getValue() == "yes"){
                possible_moves.push(next_moves["up"][0]+"7");
                // ss.getSheetByName("Executions").getRange(IS_GAME_STARTED).setValue("no");
              }
            }
          }
          catch{
            possible_moves.push(next_moves["up"]);
          }
        }
        else{
          try{
            var tem_formula = sheet.getRange(next_moves["down"]).getFormula();
            tem_formula = tem_formula.split('=IMAGE("')[1].split('")')[0];
            if (!tem_formula.endsWith('.png')){
              possible_moves.push(next_moves["down"]);
              if(ss.getSheetByName("Executions").getRange(IS_GAME_STARTED).getValue() == "yes"){
                possible_moves.push(next_moves["up"][0]+"6");
                // ss.getSheetByName("Executions").getRange(IS_GAME_STARTED).setValue("no");
              }
            }
          }
          catch{
            possible_moves.push(next_moves["down"]);
          }
        }
      }
    }
    Logger.log("MOVES");
    Logger.log(next_moves);
    Logger.log(possible_moves);
    Logger.log(red_moves);
    // Coloring possbile moves
    if(next_moves['is_pawn'] == "true"){
      ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("pawn");
      var next_moves_setr = [];
      var red_moves_setr = [];
      for(let i in possible_moves){
        if(possible_moves[i] != "" && possible_moves[i]){
          next_moves_setr.push(possible_moves[i]);
          sheet.getRange(possible_moves[i]).setBackground("yellow");
        }
      }
      var pawn_red_moves = next_moves['reds'];
      for(let j in pawn_red_moves){
        try{
          var tem_formula = sheet.getRange(pawn_red_moves[j]).getFormula();
          tem_formula = tem_formula.split('=IMAGE("')[1].split('")')[0];
          if((references[link] == "white" && references[tem_formula] == "black")||
            (references[link] == "black" && references[tem_formula] == "white")){
              sheet.getRange(pawn_red_moves[j]).setBackground("red");
              red_moves_setr.push(pawn_red_moves[j]);
          }
        }
        catch{
          //
        }
      }
      // Browser.msgBox(pawn_red_moves);
      if(next_moves_setr.length != 0){
        ss.getSheetByName("Executions").getRange(NEXT_MOVES).setValue(next_moves_setr.toString());
        ss.getSheetByName("Executions").getRange(PREV_POSITION).setValue(eaoi);
      }
      if(red_moves_setr.length != 0){
        ss.getSheetByName("Executions").getRange(RED_MOVES).setValue(red_moves_setr.toString());
      }
    }
    else{
      ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("others");
      var next_moves_setr = [];
      var red_moves_setr = [];
      for(let i in possible_moves){
        if(possible_moves[i] != "" && possible_moves[i]){
          next_moves_setr.push(possible_moves[i]);
          sheet.getRange(possible_moves[i]).setBackground("yellow");
        }
      }
      for(let i in red_moves){
        if(red_moves[i] != "" && red_moves[i]){
          red_moves_setr.push(red_moves[i]);
          sheet.getRange(red_moves[i]).setBackground("red");
        }
      }
      if(next_moves_setr.length != 0){
        ss.getSheetByName("Executions").getRange(NEXT_MOVES).setValue(next_moves_setr.toString());
        ss.getSheetByName("Executions").getRange(PREV_POSITION).setValue(eaoi);
      }
      if(red_moves_setr.length != 0){
        ss.getSheetByName("Executions").getRange(RED_MOVES).setValue(red_moves_setr.toString());
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
    var NEXT_MOVES_POS = ss.getSheetByName("Executions").getRange(NEXT_MOVES).getValue().split(",");
    var RED_MOVES_POS = ss.getSheetByName("Executions").getRange(RED_MOVES).getValue().split(",");
    if(NEXT_MOVES_POS != []){
      if(NEXT_MOVES_POS.includes(aoi)){
        ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notpawn");
        var formula = sheet.getRange('P3').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange('P3').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
      }
      else if(RED_MOVES_POS.includes(aoi)){
        ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notpawn");
        var formula = sheet.getRange('P3').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        var is_white_black = references[link];
        var next_occupying_position = get_occupying_position(is_white_black);
        var new_formula_of_opponent = sheet.getRange(aoi).getFormula();
        var new_link_of_opponent = new_formula_of_opponent.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange(next_occupying_position).setFormula('=IMAGE("' + new_link_of_opponent + '")');
        // sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
        for(let i in RED_MOVES_POS){
          sheet.getRange(RED_MOVES_POS[i]).setBackground(orig_backgrounds[RED_MOVES_POS[i]]);
        }
        sheet.getRange('P3').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
      }
      else{
        Browser.msgBox("Illegal Move!");
      }
    }
    else if(RED_MOVES != []){
      if(RED_MOVES_POS.includes(aoi)){
        ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notpawn");
        var formula = sheet.getRange('P3').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        var is_white_black = references[link];
        var next_occupying_position = get_occupying_position(is_white_black);
        var new_formula_of_opponent = sheet.getRange(aoi).getFormula();
        var new_link_of_opponent = new_formula_of_opponent.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange(next_occupying_position).setFormula('=IMAGE("' + new_link_of_opponent + '")');
        // sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
        for(let i in RED_MOVES_POS){
          sheet.getRange(RED_MOVES_POS[i]).setBackground(orig_backgrounds[RED_MOVES_POS[i]]);
        }
        sheet.getRange('P3').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
      }
    }
    else{
      Browser.msgBox("No Possbile Moves!");
    }
    ss.getSheetByName("Executions").getRange(IS_GAME_STARTED).setValue("no");
  }
  else if(is_pawn == "others"){
    var NEXT_MOVES_POS = ss.getSheetByName("Executions").getRange(NEXT_MOVES).getValue().split(",");
    var RED_MOVES_POS = ss.getSheetByName("Executions").getRange(RED_MOVES).getValue().split(",");
    if(NEXT_MOVES_POS != []){
      if(NEXT_MOVES_POS.includes(aoi)){
        ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notothers");
        var formula = sheet.getRange('P3').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange('P3').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
        for(let i in RED_MOVES_POS){
          sheet.getRange(RED_MOVES_POS[i]).setBackground(orig_backgrounds[RED_MOVES_POS[i]]);
        }
      }
      else if(RED_MOVES_POS.includes(aoi)){
        ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notothers");
        var formula = sheet.getRange('P3').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        var is_white_black = references[link];
        var next_occupying_position = get_occupying_position(is_white_black);
        var new_formula_of_opponent = sheet.getRange(aoi).getFormula();
        var new_link_of_opponent = new_formula_of_opponent.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange(next_occupying_position).setFormula('=IMAGE("' + new_link_of_opponent + '")');
        // sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
        for(let i in RED_MOVES_POS){
          sheet.getRange(RED_MOVES_POS[i]).setBackground(orig_backgrounds[RED_MOVES_POS[i]]);
        }
        sheet.getRange('P3').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
      }
      else{
        Browser.msgBox("Illegal Move!");
      }
    }
    else if(RED_MOVES != []){
      if(RED_MOVES_POS.includes(aoi)){
        ss.getSheetByName("Executions").getRange(IS_PAWN).setValue("notothers");
        var formula = sheet.getRange('P3').getFormula();
        var link = formula.split('=IMAGE("')[1].split('")')[0];
        var is_white_black = references[link];
        var next_occupying_position = get_occupying_position(is_white_black);
        var new_formula_of_opponent = sheet.getRange(aoi).getFormula();
        var new_link_of_opponent = new_formula_of_opponent.split('=IMAGE("')[1].split('")')[0];
        sheet.getRange(next_occupying_position).setFormula('=IMAGE("' + new_link_of_opponent + '")');
        // sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
        //setting original colors
        for(let i in NEXT_MOVES_POS){
          sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
        }
        for(let i in RED_MOVES_POS){
          sheet.getRange(RED_MOVES_POS[i]).setBackground(orig_backgrounds[RED_MOVES_POS[i]]);
        }
        sheet.getRange('P3').setFormula('=IMAGE("")');
        sheet.getRange(aoi).setFormula('=IMAGE("' + link + '")');
      }
    }
    else{
      Browser.msgBox("No Possbile Moves!");
    }
  }
  // sendMail();
}

function cancel(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eaoi = ss.getSheetByName("Executions").getRange(PREV_POSITION).getValue();
  var sheet = ss.getSheetByName("Chess");
  var formula = sheet.getRange('P3').getFormula();
  try{
    var link = formula.split('=IMAGE("')[1].split('")')[0];
    sheet.getRange('P3').setFormula('=IMAGE("")');
    sheet.getRange(eaoi).setFormula('=IMAGE("' + link + '")');
    var NEXT_MOVES_POS = ss.getSheetByName("Executions").getRange(NEXT_MOVES).getValue().split(",");
    var RED_MOVES_POS = ss.getSheetByName("Executions").getRange(RED_MOVES).getValue().split(",");
    for(let i in NEXT_MOVES_POS){
      sheet.getRange(NEXT_MOVES_POS[i]).setBackground(orig_backgrounds[NEXT_MOVES_POS[i]]);
    }
    for(let i in RED_MOVES_POS){
      sheet.getRange(RED_MOVES_POS[i]).setBackground(orig_backgrounds[RED_MOVES_POS[i]]);
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Chess");
  dataSheet.getRange('A16').setValue(emails.white[0][0]).setFontSize(12).setFontColor("white").setFontWeight("bold");
  dataSheet.getRange('D16').setValue(emails.black[0][0]).setFontSize(12).setFontColor("white").setFontWeight('bold');
  Browser.msgBox("Let's Start.");
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
  setEmails();
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







