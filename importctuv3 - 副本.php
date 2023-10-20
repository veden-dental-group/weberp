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
                  $allRow = 1000; 
                  //$exceldate= floatval($sheet->getCell('A3')->getValue()); //A3 放日期  取出的日期為1900-01-1到那一天的天數
                  //$tdate = date('Y-m-d', ($exceldate-25569)*86400); //25569 為1900-01-01和1970-01-01的相差天數 
                  $tdate=$file;
                  
                  $f6   = strtoupper($data->sheets[0]['cells'][6][6]); //p6 放 DESCRIPTION
                  $g6   = strtoupper($data->sheets[0]['cells'][6][7]); //O6 放 DESCRIPTION      
                  
                  $m6   = strtoupper($data->sheets[0]['cells'][6][13]); //O6 放 TOTAL 
                  $n6   = strtoupper($data->sheets[0]['cells'][6][14]); //O6 放 TOTAL           
                  $o6   = strtoupper($data->sheets[0]['cells'][6][15]); //O6 放 TOTAL 
                  $p6   = strtoupper($data->sheets[0]['cells'][6][16]); //p6 放 TOTAL   
                  $q6   = strtoupper($data->sheets[0]['cells'][6][17]); //O6 放 TOTAL      
                  
                  
                  $m7   = strtoupper($data->sheets[0]['cells'][7][13]); //O6 放 TOTAL 
                  $n7   = strtoupper($data->sheets[0]['cells'][7][14]); //O6 放 TOTAL           
                  $o7   = strtoupper($data->sheets[0]['cells'][7][15]); //O6 放 TOTAL 
                  $p7   = strtoupper($data->sheets[0]['cells'][7][16]); //p6 放 TOTAL   
                  $q7   = strtoupper($data->sheets[0]['cells'][7][17]); //O6 放 TOTAL  
                     
                    
                  if ($m6="TOTAL" or $n6="TOTAL" or $o6="TOTAL" or $p6="TOTAL" or $q6="TOTAL" )  {
                      $start=7 ;                   
                  } else if ($m6="TOTAL" or $n6="TOTAL" or $o6="TOTAL" or $p6="TOTAL" or $q6="TOTAL" )  { 
                      $start=8;
                  }
                  $iii=0;           
                  for($y = $start ; $y<=$allRow; $y++){       //由第7行開始讀資料  
                      if (($o6=='TOTAL') or ($o7=='TOTAL')) { //AA
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][5];     
                          $pqty   = $data->sheets[0]['cells'][$y][6];   
                          $pprice = $data->sheets[0]['cells'][$y][7];  
                          $ptotal = $data->sheets[0]['cells'][$y][8]; 
                          $mname  = $data->sheets[0]['cells'][$y][9]; 
                          $mqty   = $data->sheets[0]['cells'][$y][10];  
                          $mprice = $data->sheets[0]['cells'][$y][11];  
                          $mtotal = $data->sheets[0]['cells'][$y][12];  
                          $memo   = 'o6,o7';
                          $total  = $data->sheets[0]['cells'][$y][13];
                      } else if (($p6=='TOTAL') or ($p7=='TOTAL')) {     
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][6];     
                          $pqty   = $data->sheets[0]['cells'][$y][7];   
                          $pprice = $data->sheets[0]['cells'][$y][8];  
                          $ptotal = $data->sheets[0]['cells'][$y][9]; 
                          $mname  = $data->sheets[0]['cells'][$y][10]; 
                          $mqty   = $data->sheets[0]['cells'][$y][11];  
                          $mprice = $data->sheets[0]['cells'][$y][12];  
                          $mtotal = $data->sheets[0]['cells'][$y][13];  
                          $memo   = 'o6,o7';
                          $total  = $data->sheets[0]['cells'][$y][16];      
                      } else if (($q6=='TOTAL') or ($q7=='TOTAL')) {     
                          if ($f6=='DESCRIPTION') {             
                              $rx     = $data->sheets[0]['cells'][$y][1];     
                              $pname  = $data->sheets[0]['cells'][$y][6];     
                              $pqty   = $data->sheets[0]['cells'][$y][7];   
                              $pprice = $data->sheets[0]['cells'][$y][8];  
                              $ptotal = $data->sheets[0]['cells'][$y][9]; 
                              $mname  = $data->sheets[0]['cells'][$y][10]; 
                              $mqty   = $data->sheets[0]['cells'][$y][11];  
                              $mprice = $data->sheets[0]['cells'][$y][12];  
                              $mtotal = $data->sheets[0]['cells'][$y][13];  
                              $memo   = 'f6';
                              $total  = $data->sheets[0]['cells'][$y][17];       
                          } else {
                            
                              $rx     = $data->sheets[0]['cells'][$y][1];     
                              $pname  = $data->sheets[0]['cells'][$y][7];     
                              $pqty   = $data->sheets[0]['cells'][$y][8];   
                              $pprice = $data->sheets[0]['cells'][$y][9];  
                              $ptotal = $data->sheets[0]['cells'][$y][10]; 
                              $mname  = $data->sheets[0]['cells'][$y][11]; 
                              $mqty   = $data->sheets[0]['cells'][$y][12];  
                              $mprice = $data->sheets[0]['cells'][$y][13];  
                              $mtotal = $data->sheets[0]['cells'][$y][14];  
                              $memo   = 'q6';
                              $total  = $data->sheets[0]['cells'][$y][17];     
                          } 
                      } else if (($n6=='TOTAL') or ($n7=='TOTAL')) {      
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][4];     
                          $pqty   = $data->sheets[0]['cells'][$y][5];   
                          $pprice = $data->sheets[0]['cells'][$y][6];  
                          $ptotal = $data->sheets[0]['cells'][$y][7]; 
                          $mname  = $data->sheets[0]['cells'][$y][8]; 
                          $mqty   = $data->sheets[0]['cells'][$y][9];  
                          $mprice = $data->sheets[0]['cells'][$y][10];  
                          $mtotal = $data->sheets[0]['cells'][$y][11];  
                          $memo   = 'n6,n7';
                          $total  = $data->sheets[0]['cells'][$y][14]; 

                      } else if (($m6=='TOTAL') or ($m7=='TOTAL')) {        
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][3];     
                          $pqty   = $data->sheets[0]['cells'][$y][4];   
                          $pprice = $data->sheets[0]['cells'][$y][5];  
                          $ptotal = $data->sheets[0]['cells'][$y][6]; 
                          $mname  = $data->sheets[0]['cells'][$y][7]; 
                          $mqty   = $data->sheets[0]['cells'][$y][8];  
                          $mprice = $data->sheets[0]['cells'][$y][9];  
                          $mtotal = $data->sheets[0]['cells'][$y][10];  
                          $memo   = 'm6,m7';
                          $total  = $data->sheets[0]['cells'][$y][13];  
                                                                               
                      } else {
                          $msg.="$filename 格式不合!!";
                      }
                      //$memo='';
                      //寫入table中
                      if ($rx!='') {
                          $queryus = "insert into trade ( fname, tdate, cname, rx, pname, pqty, pprice, ptotal, ptotal1, mname, mqty, mprice, mtotal, mtotal1, total, memo ) values (
                                      '" . safetext($filename)        . "',  
                                      '" . $tdate                     . "',       
                                      '" . $customer                  . "',   
                                      '" . $rx                        . "',    
                                      '" . $pname                     . "',    
                                       " . floatval($pqty)            . ", 
                                       " . floatval($pprice)          . ",    
                                       " . floatval($ptotal)          . ",  
                                      '" . $ptotal                    . "',   
                                      '" . $mname                     . "',    
                                       " . floatval($mqty)            . ",    
                                       " . floatval($mprice)          . ",    
                                       " . floatval($mtotal)          . ", 
                                      '" . $mtotal                    . "',    
                                       " . floatval($total)           . ",      
                                      '" . safetext($memo)            . "')";     
                          $resultus = mysql_query($queryus) or die ('17 Trade added error!!.' .mysql_error());  
                          $iii++;
                      } 
                      
                  }     
                  echo ':'.$iii . "   ";
//                  $PHPReader->disconnectWorksheets();
                  unset($PHPReader);
              }
            }
        } 
    }  
} 
$msg.='轉檔完畢'; 
msg($msg);

?>
