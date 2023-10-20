<?php                   

//檢查有無新上傳的檔案 轉成email 通知相關人員
session_start();        
include("_data.php");

require_once 'Excel/reader.php';



//$data->setOutputEncoding('UTF8');

$arg1=$argv[1];
echo $arg1;
$dir1='d:/ctu/'.$arg1;
echo $dir2 . '  ';

//$dir1='d:/ctu/02';

$dirarray=scandir($dir1);

$msg='開始轉檔';
foreach ($dirarray as $dir2) {
    if ($dir2 != "." && $dir2 != "..") {
        $subdir=$dir1.'/'.$dir2;
        echo $subdir.'  ';    
        if (is_dir($subdir)) {           //只處理第二層目錄下的檔案
            $filearray = scandir($subdir);
            foreach($filearray as $file) {
              if ($file !='.' && $file !='..' && !(is_dir($subdir.'/'.$file))) {
                  
                  $filename=$subdir.'/'.$file;
                  $customer=$dir2;          
                  echo $filename;  
                  $data = new Spreadsheet_Excel_Reader();
                  $data->read($filename); 

                  $m6   = strtoupper($data->sheets[0]['cells'][6][13]); //O6 放 TOTAL 
                  $n6   = strtoupper($data->sheets[0]['cells'][6][14]); //O6 放 TOTAL           
                  $o6   = strtoupper($data->sheets[0]['cells'][6][15]); //O6 放 TOTAL 
                  $p6   = strtoupper($data->sheets[0]['cells'][6][16]); //p6 放 TOTAL   
                  $q6   = strtoupper($data->sheets[0]['cells'][6][17]); //O6 放 TOTAL    
                  $sname6= '1'.strtoupper($data->sheets[0]['cells'][6][1]) . ' 2'. strtoupper($data->sheets[0]['cells'][6][2])  .
                           ' 3'.strtoupper($data->sheets[0]['cells'][6][3]) . ' 4'.strtoupper($data->sheets[0]['cells'][6][4])  . 
                           ' 5'.strtoupper($data->sheets[0]['cells'][6][5]) . ' 6'.strtoupper($data->sheets[0]['cells'][6][6])  . 
                           ' 7'.strtoupper($data->sheets[0]['cells'][6][7]) . ' 8'.strtoupper($data->sheets[0]['cells'][6][8])  . 
                           ' 9'.strtoupper($data->sheets[0]['cells'][6][9]) . ' 10'.strtoupper($data->sheets[0]['cells'][6][10])  . 
                           ' 11'.strtoupper($data->sheets[0]['cells'][6][11]) . ' 12'.strtoupper($data->sheets[0]['cells'][6][12])  . 
                           ' 13'.strtoupper($data->sheets[0]['cells'][6][13]) . ' 14'.strtoupper($data->sheets[0]['cells'][6][14])  . 
                           ' 15'.strtoupper($data->sheets[0]['cells'][6][15]) . ' 16'.strtoupper($data->sheets[0]['cells'][6][16])  . 
                           ' 17'.strtoupper($data->sheets[0]['cells'][6][17]) . ' 18'.strtoupper($data->sheets[0]['cells'][6][18])  . 
                           ' 19'.strtoupper($data->sheets[0]['cells'][6][19]) . ' 20'.strtoupper($data->sheets[0]['cells'][6][20])  . 
                           ' 21'.strtoupper($data->sheets[0]['cells'][6][21]) . ' 22'.strtoupper($data->sheets[0]['cells'][6][22])  ; 
                  
                  $sname7= '1'.strtoupper($data->sheets[0]['cells'][7][1]) . ' 2'. strtoupper($data->sheets[0]['cells'][7][2])  .
                           ' 3'.strtoupper($data->sheets[0]['cells'][7][3]) . ' 4'.strtoupper($data->sheets[0]['cells'][7][4])  . 
                           ' 5'.strtoupper($data->sheets[0]['cells'][7][5]) . ' 6'.strtoupper($data->sheets[0]['cells'][7][6])  . 
                           ' 7'.strtoupper($data->sheets[0]['cells'][7][7]) . ' 8'.strtoupper($data->sheets[0]['cells'][7][8])  . 
                           ' 9'.strtoupper($data->sheets[0]['cells'][7][9]) . ' 10'.strtoupper($data->sheets[0]['cells'][7][10])  . 
                           ' 11'.strtoupper($data->sheets[0]['cells'][7][11]) . ' 12'.strtoupper($data->sheets[0]['cells'][7][12])  . 
                           ' 13'.strtoupper($data->sheets[0]['cells'][7][13]) . ' 14'.strtoupper($data->sheets[0]['cells'][7][14])  . 
                           ' 15'.strtoupper($data->sheets[0]['cells'][7][15]) . ' 16'.strtoupper($data->sheets[0]['cells'][7][16])  . 
                           ' 17'.strtoupper($data->sheets[0]['cells'][7][17]) . ' 18'.strtoupper($data->sheets[0]['cells'][7][18])  . 
                           ' 19'.strtoupper($data->sheets[0]['cells'][7][19]) . ' 20'.strtoupper($data->sheets[0]['cells'][7][20])  . 
                           ' 21'.strtoupper($data->sheets[0]['cells'][7][21]) . ' 22'.strtoupper($data->sheets[0]['cells'][7][22])  ; 
                  


                          $queryus = "insert into style1 ( sname,sfilename, sline ) values (
                                      '" . $sname6            . "',  
                                      '" . $filename          . "',     
                                      '" . 6                  . "')";     
                          $resultus = mysql_query($queryus) or die ('17 Trade added error!!.' .mysql_error());  
                          
                          $queryus = "insert into style1 ( sname,sfilename, sline ) values (
                                      '" . $sname7            . "',  
                                      '" . $filename          . "',     
                                      '" . 7                  . "')";     
                          $resultus = mysql_query($queryus) or die ('17 Trade added error!!.' .mysql_error());
                    
                      
                  }     
                 // echo ':'.$iii . "   ";
//                  $PHPReader->disconnectWorksheets();
                //  unset($PHPReader);
              }
            }
        } 
    }  

$msg.='轉檔完畢'; 
msg($msg);

?>
