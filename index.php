<?php
    require 'vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    date_default_timezone_set('Asia/Kuala_Lumpur');

    function cleanString($input) {
        $string = str_split($input);
        $output = "";
        for ($i=0; $i<count($string); $i++) {
            if (mb_ord($string[$i],'UTF-16') <= 127) {
                $output .= $string[$i];
            }
        }
        return $output;
    }

    function RowNotEmpty($array){
        foreach($array as $data){
            if(null !== $data){
                return false;
            }
        }
        return true;
    }

    if($_SERVER["REQUEST_METHOD"] == "POST"){
        $targetPath = 'tmp/' . $_FILES['my_file_input']['name'];
        move_uploaded_file($_FILES['my_file_input']['tmp_name'], $targetPath);
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($targetPath);
        $names = $spreadsheet->getSheetNames();
        $edi_string = "";
        $revcode = $_POST["recv_code"];
        $callcode = $_POST["callsign_code"];
        $dateraw =  date('Ymd').":".date('hi');
        $refno =  date('Ymdhis');
        
        foreach($names as $names){
            $sheet = $spreadsheet->getSheetByName($names);
            $lastrow = $sheet->getHighestRow();
            $line = 0;
            $contcount = 0;

            $edi = "UNB+UNOA:2+KMT+".$revcode."+".$dateraw."+".$refno."'\n";$line++;
            $edi .= "UNH+".$refno."+COPRAR:D:00B:UN:SMDG21+LOADINGCOPRAR'\n";$line++;

            $tmpstring = $sheet->getcell('B2')->getValue();
            $tmp = strtotime(str_replace('/','-',$tmpstring));
            $vsdatetime = date('YmdHis',$tmp);
            $edi .= "BGM+45+".$vsdatetime."+5'\n";$line++;

            $tmp = explode("/",$sheet->getcell('D4')->getValue());
            $voyage = $tmp[0];
            $callsign = $tmp[1];
            $opr = $tmp[2];
            $vslname = $sheet->getcell('B4')->getValue();
            $edi .= "TDT+20+".$voyage."+1++172:".$opr."+++".$callcode.":103::".$vslname."'\n"; $line++;
            $edi .= "RFF+VON:".$voyage."'\n"; $line++;
            $edi .= "NAD+CA+".$opr."'\n"; $line++;

            $currow = 9;
            while($currow <= $lastrow){
                $row = $sheet->rangeToArray("A".$currow.":AB".$currow);
                if(!RowNotEmpty($row)){
                    $contcount++;
                    $rowCells = $row[0];
                    //rowCells[3] //5 - F, 4 - E
                    $fe = "5";
                    if(!is_null($rowCells[3]) && $rowCells[3]=="E"){$fe = "4";}
                    //2 TS - N, 6 TS - Y
                    $type = "2";
                    if(!is_null($rowCells[11]) && $rowCells[11]=="Y"){ $type = "6";}
                    
                    if(!is_null($rowCells[1]) && !is_null($rowCells[7])){ $edi .= "EQD+CN+".$rowCells[1]."+".$rowCells[7].":102:5++".$type."+".$fe."'\n";$line++;}
                    if(!is_null($rowCells[6])){ $edi .= "LOC+11+".$rowCells[5].":139:6'\n"; $line++;}
                    if(!is_null($rowCells[6])){ $edi .= "LOC+7+".$rowCells[6].":139:6'\n"; $line++; }
                    if(!is_null($rowCells[19])) { $edi .= "LOC+9+".$rowCells[19].":139:6'\n"; $line++; }
                    if(!is_null($rowCells[13])) { $edi .= "MEA+AAE+VGM+KGM:".$rowCells[13]."'\n"; $line++; }
                    $cell17 = $rowCells[17];
                    if(!is_null($rowCells[17]) && trim($rowCells[17])!="" && trim($rowCells[17])!="/") {
                    $tmp = explode(',',$rowCells[17]);
                    for($i=0; $i<count($tmp); $i++) {
                        $dim = explode('/',$rowCells[17]);
                        if(trim($dim[0])=="OF") {
                            $edi .= "DIM+5+CMT:".trim($dim[1])."'\n";$line++;
                        }
                        if(trim($dim[0])=="OB") {
                            $edi .= "DIM+6+CMT:".trim($dim[1])."'\n";$line++;
                        }
                        if(trim($dim[0])=="OR") {
                            $edi .= "DIM+7+CMT::".trim($dim[1])."'\n";$line++;
                        }
                        if(trim($dim[0])=="OL") {
                            $edi .= "DIM+8+CMT::".trim($dim[1])."'\n";$line++;
                        }
                        if(trim($dim[0])=="OH") {
                            $edi .= "DIM+9+CMT:::".trim($dim[1])."'\n";$line++;
                        }
                    }
                    }
                    if(!is_null($rowCells[15]) && trim($rowCells[15])!="" && trim($rowCells[15])!="/") {
                    $temperature = $rowCells[15];
                    $temperature = str_replace(" ", "",$temperature);
                    $temperature = str_replace("C", "",$temperature);
                    $temperature = str_replace("+", "",$temperature);
                    $edi .= "TMP+2+".$temperature.":CEL'\n"; $line++;
                    }
                    if(!is_null($rowCells[25]) && trim($rowCells[25])!="" && trim($rowCells[25])!="/") {
                    $tmp = explode(',',$rowCells[25]);
                    if($tmp[0]=="L") {
                        $edi .= "SEL+".$tmp[1]."+CA'\n"; $line++; //seal L - CA, S - SH, M - CU
                    }
                    if($tmp[0]=="S") {
                        $edi .= "SEL+".$tmp[1]."+SH'\n"; $line++; //seal L - CA, S - SH, M - CU
                    }
                    if($tmp[0]=="M") {
                        $edi .= "SEL+".$tmp[1]."+CU'\n"; $line++; //seal L - CA, S - SH, M - CU
                    }
                    }
                    if(!is_null($rowCells[8])){ $edi .= "FTX+AAI+++".$rowCells[8]."'\n"; $line++; }                      
                    
                    if(!is_null($rowCells[12]) && trim($rowCells[12])!="" && trim($rowCells[12])!="/") {
                    $edi .= "FTX+AAA+++".trim(cleanString($rowCells[12]))."'\n";$line++;
                    }
                    if(!is_null($rowCells[18]) && trim($rowCells[18])!="" && trim($rowCells[18])!="/") {
                    $edi .= "FTX+HAN++".$rowCells[18]."'\n";$line++;
                    }
                    if(!is_null($rowCells[14]) && trim($rowCells[14])!="" && trim($rowCells[14])!="/") {
                    $tmp = explode('/',$rowCells[14]);
                    $edi .= "DGS+IMD+".$tmp[0]."+".$tmp[1]."'\n";$line++;
                    }
                    if(!is_null($rowCells[2]) && trim($rowCells[2])!="") { $edi .= "NAD+CF+".$rowCells[2].":160:ZZZ'\n"; $line++; }

                    
                }
                $currow++;
            }

            $edi .= "CNT+16:".$contcount."'\n";$line++;
            $edi .= "UNT+".$line."+".$refno."'\n";
            $edi .= "UNZ+1+".$refno."'";

            $edi_string.=$edi;
        }
    }
?>
<html lang="en"><head>
<!-- Required meta tags -->
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
<title>Export Booking Excel to Coprar Converter</title>
<!-- Bootstrap CSS -->
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha256-4+XzXVhsDmqanXGHaHvgh1gMQKX40OUvDEBTu8JcmNs=" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js"></script>

</head>
<body>
         
<script>

          parseExcel = function(file) {
            var reader = new FileReader();

            reader.onload = function(e) {
              var data = e.target.result;
              var workbook = XLSX.read(data, {
                type: 'binary'
              });

              workbook.SheetNames.forEach(function(sheetName) {
                // Here is your object
                //var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                var XL_row_object = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
                //var json_object = JSON.stringify(XL_row_object);
                //console.log(XL_row_object)
                //console.log(json_object);
                var line = 0;
                var contcount = 0;
                var allRows = XL_row_object.split(/\r?\n|\r/);
                //var table = '<table>';
                var dt = new Date();
                var refno = get_date_str(dt, "");
                var edi = "UNB+UNOA:2+KMT+"+$("#recv_code").val()+"+"+get_date_str(dt, "daterawonly")+":"+get_date_str(dt, "timetominrawonly")+"+"+refno+"'\n"; 
                edi += "UNH+"+refno+"+COPRAR:D:00B:UN:SMDG21+LOADINGCOPRAR'\n"; line++;
                                
                //process header
                var report_dt = ""; var voyage = ""; var vslname = ""; var callsign = ""; var opr = "";
                for (var singleRow = 0; singleRow < allRows.length; singleRow++) {
                    if(singleRow>6) break;
                    let rowCells = allRows[singleRow].split(',');
                    if(singleRow==1) {
                        let tmpdt = rowCells[1].split('/');
                        let day = tmpdt[0];
                        let month = tmpdt[1];
                        let tmpyear = tmpdt[2].split(' ');
                        let report_date = new Date(tmpyear[0] + "-" + month + "-" + day + " " + tmpyear[1]);
                        report_dt = get_date_str(report_date, "");
                    }
                    if(singleRow==3) {
                        if(typeof rowCells[3]!="undefined") {
                            let tmp = rowCells[3].split('/');
                            voyage = tmp[0];
                            callsign = tmp[1];
                            opr = tmp[2];
                            vslname = rowCells[1];
                        }
                    }
                }                
                edi += "BGM+45+"+report_dt+"+5'\n"; line++;
                edi += "TDT+20+"+voyage+"+1++172:"+opr+"+++"+$("#callsign_code").val()+":103::"+vslname+"'\n"; line++;
                edi += "RFF+VON:"+voyage+"'\n"; line++;
                edi += "NAD+CA+"+opr+"'\n"; line++;
                let tmp; let dim;
                for (singleRow = 0; singleRow < allRows.length; singleRow++) {
                  if(typeof allRows[singleRow]!="undefined") {
                    //let rowCells = allRows[singleRow].split(',');
                    let rowCells = CSVtoArray(allRows[singleRow]);
                    if(singleRow>7) {
                        contcount++;
                        //rowCells[3] //5 - F, 4 - E
                        let fe = "5";
                        if(typeof rowCells[3]!="undefined" && rowCells[3]=="E") fe = "4";
                        //2 TS - N, 6 TS - Y
                        let type = "2";
                        if(typeof rowCells[11]!="undefined" && rowCells[11]=="Y") type = "6";
                        
                        if(typeof rowCells[1]!="undefined" && typeof rowCells[7]!="undefined") { edi += "EQD+CN+"+rowCells[1]+"+"+rowCells[7]+":102:5++"+type+"+"+fe+"'\n"; line++; }
                        if(typeof rowCells[6]!="undefined") { edi += "LOC+11+"+rowCells[5]+":139:6'\n"; line++; }
                        if(typeof rowCells[6]!="undefined") { edi += "LOC+7+"+rowCells[6]+":139:6'\n"; line++; }
                        if(typeof rowCells[19]!="undefined") { edi += "LOC+9+"+rowCells[19]+":139:6'\n"; line++; }
                        if(typeof rowCells[13]!="undefined") { edi += "MEA+AAE+VGM+KGM:"+rowCells[13]+"'\n"; line++; }
                        if(typeof rowCells[17]!="undefined" && $.trim(rowCells[17])!="" && $.trim(rowCells[17])!="/") {
                          tmp = rowCells[17].split(',');
                          for(let i=0; i<tmp.length; i++) {
                              dim = rowCells[17].split('/');
                              if($.trim(dim[0])=="OF") {
                                  edi += "DIM+5+CMT:"+$.trim(dim[1])+"'\n"; line++;
                              }
                              if($.trim(dim[0])=="OB") {
                                  edi += "DIM+6+CMT:"+$.trim(dim[1])+"'\n"; line++;
                              }
                              if($.trim(dim[0])=="OR") {
                                  edi += "DIM+7+CMT::"+$.trim(dim[1])+"'\n"; line++;
                              }
                              if($.trim(dim[0])=="OL") {
                                  edi += "DIM+8+CMT::"+$.trim(dim[1])+"'\n"; line++;
                              }
                              if($.trim(dim[0])=="OH") {
                                  edi += "DIM+9+CMT:::"+$.trim(dim[1])+"'\n"; line++;
                              }
                          }
                        }
                        if(typeof rowCells[15]!="undefined" && $.trim(rowCells[15])!="" && $.trim(rowCells[15])!="/") {
                          let temperature = rowCells[15];
                          temperature = temperature.replace(" ", "");
                          temperature = temperature.replace("C", "");
                          temperature = temperature.replace("+", "");
                          edi += "TMP+2+"+temperature+":CEL'\n"; line++;
                        }
                        if(typeof rowCells[25]!="undefined" && $.trim(rowCells[25])!="" && $.trim(rowCells[25])!="/") {
                          let tmp = rowCells[25].split(',');
                          if(tmp[0]=="L") {
                              edi += "SEL+"+tmp[1]+"+CA'\n"; line++; //seal L - CA, S - SH, M - CU
                          }
                          if(tmp[0]=="S") {
                              edi += "SEL+"+tmp[1]+"+SH'\n"; line++; //seal L - CA, S - SH, M - CU
                          }
                          if(tmp[0]=="M") {
                              edi += "SEL+"+tmp[1]+"+CU'\n"; line++; //seal L - CA, S - SH, M - CU
                          }
                        }
                        if(typeof rowCells[8]!="undefined") { edi += "FTX+AAI+++"+rowCells[8]+"'\n"; line++; }                      
                        
                        if(typeof rowCells[12]!="undefined" && $.trim(rowCells[12])!="" && $.trim(rowCells[12])!="/") {
                          edi += "FTX+AAA+++"+$.trim(cleanString(rowCells[12]))+"'\n"; line++;
                        }
                        if(typeof rowCells[18]!="undefined" && $.trim(rowCells[18])!="" && $.trim(rowCells[18])!="/") {
                          edi += "FTX+HAN++"+rowCells[18]+"'\n"; line++;
                        }
                        if(typeof rowCells[14]!="undefined" && rowCells[14]!="" && $.trim(rowCells[14])!="/") {
                          tmp = rowCells[14].split('/');
                          edi += "DGS+IMD+"+tmp[0]+"+"+tmp[1]+"'\n"; line++;
                        }
                        if(typeof rowCells[2]!="undefined" && $.trim(rowCells[2])!="") { edi += "NAD+CF+"+rowCells[2]+":160:ZZZ'\n"; line++; } //box 
                        //if(opr!="") { edi += "NAD+CA+"+opr+":160:ZZZ'\n"; line++; } //vsl
                        //if(typeof rowCells[27]!="undefined" && $.trim(rowCells[27])!="")  { edi += "NAD+GF+"+rowCells[27]+":160:ZZZ'\n"; line++; } //slot
                    }                    

                    /*if (singleRow === 0) {
                      table += '<thead>';
                      table += '<tr>';
                    } else {
                      table += '<tr>';
                    }
                    var rowCells = allRows[singleRow].split(',');
                    for (let rowCell = 0; rowCell < rowCells.length; rowCell++) {
                      if (singleRow === 0) {
                        table += '<th>';
                        table += rowCells[rowCell];
                        table += '</th>';
                      } else {
                        table += '<td>';
                        table += rowCells[rowCell];
                        table += '</td>';
                      }
                    }
                    if (singleRow === 0) {
                      table += '</tr>';
                      table += '</thead>';
                      table += '<tbody>';
                    } else {
                      table += '</tr>';
                    }*/
                  }
                } 
                contcount--;
                edi += "CNT+16:"+contcount+"'\n"; line++; line++;
                edi += "UNT+"+line+"+"+refno+"'\n";
                edi += "UNZ+1+"+refno+"'";
                //table += '</tbody>';
                //table += '</table>';
                $('#my_file_output').val(edi);

              })

            };

            reader.onerror = function(ex) {
              console.log(ex);
            };

            reader.readAsBinaryString(file);
          };
        
        var oFileIn;

        $(function() {
            oFileIn = document.getElementById('my_file_input');
            if(oFileIn.addEventListener) {
                oFileIn.addEventListener('change', filePicked, false);
            }
        });


        function filePicked(oEvent) {
            // Get The File From The Input
            var oFile = oEvent.target.files[0];
            var sFilename = oFile.name;
            parseExcel(oFile)
            
        }
        
        function copy() {
            /* Get the text field */
            var copyText = document.getElementById("ediholder");

            /* Select the text field */
            copyText.select();
            copyText.setSelectionRange(0, 99999); /* For mobile devices */

            /* Copy the text inside the text field */
            document.execCommand("copy");

            /* Alert the copied text */
            alert("Text Copied!");
        } 
        
        function get_date_str(d, type) {
            var now = d;
            var dt = now.getDate();
            dt = (String(dt).length<2)? String("0") + String(dt) : dt;
            var hrs = now.getHours();
            hrs = (String(hrs).length<2)? String("0") + String(hrs) : hrs;
            var min = now.getMinutes();
            min = (String(min).length<2)? String("0") + String(min) : min;
            var sec = now.getSeconds();
            sec = (String(sec).length<2)? String("0") + String(sec) : sec;
            var mth = (now.getMonth() + 1);
            mth = (String(mth).length<2)? String("0") + String(mth) : mth;
            if(type=="daterawonly") {
                return now.getFullYear()+''+String(mth)+''+String(dt);
            } else if(type=="timetominrawonly"){
                return String(hrs)+''+String(min);
            } else {
                return now.getFullYear()+''+String(mth)+''+String(dt)+''+String(hrs)+''+String(min)+''+String(sec);
            }
            //return now.getHours()+':'+String(min)+':'+String(sec);
        }
        
        function CSVtoArray(text) {
            var re_valid = /^\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*(?:,\s*(?:'[^'\\]*(?:\\[\S\s][^'\\]*)*'|"[^"\\]*(?:\\[\S\s][^"\\]*)*"|[^,'"\s\\]*(?:\s+[^,'"\s\\]+)*)\s*)*$/;
            var re_value = /(?!\s*$)\s*(?:'([^'\\]*(?:\\[\S\s][^'\\]*)*)'|"([^"\\]*(?:\\[\S\s][^"\\]*)*)"|([^,'"\s\\]*(?:\s+[^,'"\s\\]+)*))\s*(?:,|$)/g;

            // Return NULL if input string is not well formed CSV string.
            if (!re_valid.test(text)) return null;

            var a = []; // Initialize array to receive values.
            text.replace(re_value, // "Walk" the string using replace with callback.
                function(m0, m1, m2, m3) {

                    // Remove backslash from \' in single quoted values.
                    if (m1 !== undefined) a.push(m1.replace(/\\'/g, "'"));

                    // Remove backslash from \" in double quoted values.
                    else if (m2 !== undefined) a.push(m2.replace(/\\"/g, '"'));
                    else if (m3 !== undefined) a.push(m3);
                    return ''; // Return empty string.
                });

            // Handle special case of empty last value.
            if (/,\s*$/.test(text)) a.push('');
            return a;
        };
        
        function cleanString(input) {
            var output = "";
            for (var i=0; i<input.length; i++) {
                if (input.charCodeAt(i) <= 127) {
                    output += input.charAt(i);
                }
            }
            return output;
        }
</script>
<div class="container">
<div class="card">
    <div class="card-body">
        <h5 class="card-title">Export Booking Excel to Coprar Converter</h5>
        <?php
            if(isset($test)){
                echo 'The sheet name is '.$test;
            }
        ?>
        <form action="<?php echo htmlspecialchars($_SERVER["PHP_SELF"]); ?>" method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label for="recv_code">Receiver Code:</label><input class="form-control" type="text" id="recv_code" name="recv_code" value="RECEIVER" required>
                <p><small>Please change before file select.</small></p>
            </div>
            <div class="form-group">
                <label for="recv_code">Callsign Code:</label><input class="form-control" type="text" id="callsign_code" name="callsign_code" value="XXXXX" required>
                <p><small>Please change before file select.</small></p>
            </div>
            <div class="form-group">
                <label for="my_file_input">Export booking excel file:</label><input class="form-control" type="file" id="my_file_input" name="my_file_input" required>
                <p><small><a href="sample.xlsx">Sample Excel</a></small></p>
            </div>
            <button type="submit">Submit</button>
        </form><br>
        <div><textarea class="form-control" rows="20" cols="40" id="my_file_output"><?php 
                if(isset($edi_string)){
                    echo $edi_string;
                }
            ?></textarea></div>
    </div>
</div>
</div>


</body></html>

