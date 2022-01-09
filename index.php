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

</head>
<body>
         
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
        <div><textarea class="form-control" rows="20" cols="40" id="my_file_output" disabled><?php 
                if(isset($edi_string)){
                    echo $edi_string;
                }
            ?></textarea></div>
    </div>
</div>
</div>


</body></html>

