<?php                   

//檢查有無新上傳的檔案 轉成email 通知相關人員
session_start();        
include("_data.php");

require_once 'excel_reader2.php';           

$arg1=$argv[1];
echo $arg1;
$dir1='d:/ctu/'.$arg1;
echo $dir2 . '  ';

$dir1='d:/ctu/14';

$dirarray=scandir($dir1);

$msg='';
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
                  
                  $data = new Spreadsheet_Excel_Reader($filename);
                  $data->setOutputEncoding('CP936');  
                  $data->dump(true,true);
                  $tdate=$file;
                  
                  $d74  = strtoupper($data->sheets[0]['cells'][7][4]); //74 放 DESCRIPTION
                  $d75  = strtoupper($data->sheets[0]['cells'][7][5]); //75 放 DESCRIPTION      

                  $d63  = strtoupper($data->sheets[0]['cells'][6][3]); //75 放 DESCRIPTION 
                  $d64  = strtoupper($data->sheets[0]['cells'][6][4]); //75 放 DESCRIPTION  
                  $d65  = strtoupper($data->sheets[0]['cells'][6][5]); //75 放 DESCRIPTION  
                  $d66  = strtoupper($data->sheets[0]['cells'][6][6]); //75 放 DESCRIPTION  
                  $d67  = strtoupper($data->sheets[0]['cells'][6][7]); //75 放 DESCRIPTION   
                  
                  $d617  = strtoupper($data->sheets[0]['cells'][6][17]); //75 放 TOTAL
                  $d620  = strtoupper($data->sheets[0]['cells'][6][20]); //75 放 TOTAL  
                     
                    
                  if ($d74=='DESCRIPTION' or $d75=='DESCRIPTION')  {
                      $start=8 ;                   
                  } else { 
                      $start=7;
                  }
                  $iii=0;           
                  for($y = $start ; $y<=$allRow; $y++){       //由第7行開始讀資料  
                      if ($d74=='DESCRIPTION') {
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][4];     
                          $pqty   = $data->sheets[0]['cells'][$y][5];   
                          $pprice = $data->sheets[0]['cells'][$y][6];  
                          $ptotal = $data->sheets[0]['cells'][$y][7]; 
                          $mname  = $data->sheets[0]['cells'][$y][8]; 
                          $mqty   = $data->sheets[0]['cells'][$y][9];  
                          $mprice = $data->sheets[0]['cells'][$y][10];  
                          $mtotal = $data->sheets[0]['cells'][$y][11];  
                          $memo   = 'd74';
                          $total  = $data->sheets[0]['cells'][$y][14];
                      } else if ($d75=='DESCRIPTION') {     
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][5];     
                          $pqty   = $data->sheets[0]['cells'][$y][6];   
                          $pprice = $data->sheets[0]['cells'][$y][7];  
                          $ptotal = $data->sheets[0]['cells'][$y][8]; 
                          $mname  = $data->sheets[0]['cells'][$y][9]; 
                          $mqty   = $data->sheets[0]['cells'][$y][10];  
                          $mprice = $data->sheets[0]['cells'][$y][11];  
                          $mtotal = $data->sheets[0]['cells'][$y][12];  
                          $memo   = 'd75';
                          $total  = $data->sheets[0]['cells'][$y][15]; 
                      } else if ($d63=='DESCRIPTION') {     
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][3];     
                          $pqty   = $data->sheets[0]['cells'][$y][4];   
                          $pprice = $data->sheets[0]['cells'][$y][5];  
                          $ptotal = $data->sheets[0]['cells'][$y][6]; 
                          $mname  = $data->sheets[0]['cells'][$y][7]; 
                          $mqty   = $data->sheets[0]['cells'][$y][8];  
                          $mprice = $data->sheets[0]['cells'][$y][9];  
                          $mtotal = $data->sheets[0]['cells'][$y][10];  
                          $memo   = 'd63';
                          $total  = $data->sheets[0]['cells'][$y][13];  
                      } else if ($d64=='DESCRIPTION') {     
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][4];     
                          $pqty   = $data->sheets[0]['cells'][$y][5];   
                          $pprice = $data->sheets[0]['cells'][$y][6];  
                          $ptotal = $data->sheets[0]['cells'][$y][7]; 
                          $mname  = $data->sheets[0]['cells'][$y][8]; 
                          $mqty   = $data->sheets[0]['cells'][$y][9];  
                          $mprice = $data->sheets[0]['cells'][$y][10];  
                          $mtotal = $data->sheets[0]['cells'][$y][11];  
                          $memo   = 'd64';
                          $total  = $data->sheets[0]['cells'][$y][14];  
                      } else if ($d65=='DESCRIPTION') {     
                          $rx     = $data->sheets[0]['cells'][$y][1];     
                          $pname  = $data->sheets[0]['cells'][$y][5];     
                          $pqty   = $data->sheets[0]['cells'][$y][6];   
                          $pprice = $data->sheets[0]['cells'][$y][7];  
                          $ptotal = $data->sheets[0]['cells'][$y][8]; 
                          $mname  = $data->sheets[0]['cells'][$y][9]; 
                          $mqty   = $data->sheets[0]['cells'][$y][10];  
                          $mprice = $data->sheets[0]['cells'][$y][11];  
                          $mtotal = $data->sheets[0]['cells'][$y][12];  
                          $memo   = 'd65';
                          $total  = $data->sheets[0]['cells'][$y][15];  
                      } else if ($d66=='DESCRIPTION') {     
                          if ($d617=='TOTAL'){
                              $rx     = $data->sheets[0]['cells'][$y][1];     
                              $pname  = $data->sheets[0]['cells'][$y][6];     
                              $pqty   = $data->sheets[0]['cells'][$y][7];   
                              $pprice = $data->sheets[0]['cells'][$y][8];  
                              $ptotal = $data->sheets[0]['cells'][$y][9]; 
                              $mname  = $data->sheets[0]['cells'][$y][10]; 
                              $mqty   = $data->sheets[0]['cells'][$y][11];  
                              $mprice = $data->sheets[0]['cells'][$y][12];  
                              $mtotal = $data->sheets[0]['cells'][$y][13];  
                              $memo   = 'd66';
                              $total  = $data->sheets[0]['cells'][$y][17]; 
                          } else {
                              $rx     = $data->sheets[0]['cells'][$y][1];     
                              $pname  = $data->sheets[0]['cells'][$y][6];     
                              $pqty   = $data->sheets[0]['cells'][$y][7];   
                              $pprice = $data->sheets[0]['cells'][$y][8];  
                              $ptotal = $data->sheets[0]['cells'][$y][9]; 
                              $mname  = $data->sheets[0]['cells'][$y][10]; 
                              $mqty   = $data->sheets[0]['cells'][$y][11];  
                              $mprice = $data->sheets[0]['cells'][$y][12];  
                              $mtotal = $data->sheets[0]['cells'][$y][13];  
                              $memo   = 'd66';
                              $total  = $data->sheets[0]['cells'][$y][16];  
                          }
                      } else if ($d67=='DESCRIPTION') {     
                          if ($d620=='TOTAL'){    
                              $rx     = $data->sheets[0]['cells'][$y][1];     
                              $pname  = $data->sheets[0]['cells'][$y][7];     
                              $pqty   = $data->sheets[0]['cells'][$y][8];   
                              $pprice = $data->sheets[0]['cells'][$y][9];  
                              $ptotal = $data->sheets[0]['cells'][$y][10]; 
                              $mname  = $data->sheets[0]['cells'][$y][11]; 
                              $mqty   = $data->sheets[0]['cells'][$y][12];  
                              $mprice = $data->sheets[0]['cells'][$y][13];  
                              $mtotal = $data->sheets[0]['cells'][$y][14];  
                              $memo   = 'd67';
                              $total  = $data->sheets[0]['cells'][$y][20]; 
                            
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
                              $memo   = 'd67';
                              $total  = $data->sheets[0]['cells'][$y][17];    
                          } 
                      } else {
                          $msg.="$filename unknown format!!  ";
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
                          commit;
                          $iii++;
                      } 
                      
                  }     
                  echo ':'.$iii . "   ";   
              }
            }
        } 
    }  
} 
echo $msg;

?>
