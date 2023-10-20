<?php           

//使用phpexcel .xlsx
//        


session_start();        
include("_data.php");

require_once 'classes/PHPExcel.php'; 
require_once 'classes/PHPExcel/IOFactory.php';  

$arg1=$argv[1];
echo $arg1;
$dir1='d:/new3/'.$arg1;
echo $dir2;

//$dir1='d:/ctu/14';

$dirarray=scandir($dir1);
//$edate=date('Y-m-d');
$msg='';
foreach ($dirarray as $dir2) {
    if ($dir2 != "." && $dir2 != "..") {
        $subdir=$dir1.'/'.$dir2;
        if (is_dir($subdir)) {           //只處理第二層目錄下的檔案
            $filearray = scandir($subdir);
            foreach($filearray as $file) {
              if ($file !='.' && $file !='..' && !(is_dir($subdir.'/'.$file))) {
                  $filename=$subdir.'/'.$file;
                  $customer=$dir2;          
                  echo $filename;
                  sleep(1);  
                  $PHPReader = new PHPExcel_Reader_Excel2007();                                                
                  $PHPExcel = $PHPReader->load( $filename);   
                  sleep(1);  
                  $sheet = $PHPExcel->getActiveSheet(0);      
                  $allRow = 1000;
                  $exceldate= floatval($sheet->getCell('A3')->getValue()); //A3 放日期  取出的日期為1900-01-1到那一天的天數
                  $tdate = date('Y-m-d', ($exceldate-25569)*86400); //25569 為1900-01-01和1970-01-01的相差天數 
                  
                  $d7  = strtoupper($sheet->getCell('D7')->getValue()); //74 放 DESCRIPTION
                  $e7  = strtoupper($sheet->getCell('E7')->getValue()); //75 放 DESCRIPTION      

                  $c6  = strtoupper($sheet->getCell('C6')->getValue()); //75 放 DESCRIPTION 
                  $d6  = strtoupper($sheet->getCell('D6')->getValue()); //75 放 DESCRIPTION  
                  $e6  = strtoupper($sheet->getCell('E6')->getValue()); //75 放 DESCRIPTION  
                  $f6  = strtoupper($sheet->getCell('F6')->getValue()); //75 放 DESCRIPTION  
                  $g6  = strtoupper($sheet->getCell('G6')->getValue()); //75 放 DESCRIPTION   
                  
                  $n6  = strtoupper($sheet->getCell('N6')->getValue()); //75 放 TOTAL 
                  $q6  = strtoupper($sheet->getCell('Q6')->getValue()); //75 放 TOTAL
                  $t6  = strtoupper($sheet->getCell('T6')->getValue()); //75 放 TOTAL         
                     
                    
                  if ($d7=='DESCRIPTION' or $e7=='DESCRIPTION')  {
                      $start=8 ;                   
                  } else { 
                      $start=7;
                  }
                  $iii=0;         
                  for($y = $start ; $y<=$allRow; $y++){       //由第7行開始讀資料       
                  
                      if ($d7=='DESCRIPTION') {
                          $rx     = $sheet->getCell('A'.$y)->getValue();     
                          $pname  = $sheet->getCell('D'.$y)->getValue();     
                          $pqty   = $sheet->getCell('E'.$y)->getValue();   
                          $pprice = $sheet->getCell('F'.$y)->getValue(); 
                          $ptotal = $sheet->getCell('G'.$y)->getCalculatedValue(); 
                          $mname  = $sheet->getCell('H'.$y)->getValue(); 
                          $mqty   = $sheet->getCell('I'.$y)->getValue(); 
                          $mprice = $sheet->getCell('J'.$y)->getValue(); 
                          $mtotal = $sheet->getCell('K'.$y)->getCalculatedValue(); 
                          $memo   = 'd7';
                          $total  = $sheet->getCell('N'.$y)->getCalculatedValue(); 
                      } else if ($e7=='DESCRIPTION') {     
                          $rx     = $sheet->getCell('A'.$y)->getValue();     
                          $pname  = $sheet->getCell('E'.$y)->getValue();     
                          $pqty   = $sheet->getCell('F'.$y)->getValue();   
                          $pprice = $sheet->getCell('G'.$y)->getValue(); 
                          $ptotal = $sheet->getCell('H'.$y)->getCalculatedValue(); 
                          $mname  = $sheet->getCell('I'.$y)->getValue(); 
                          $mqty   = $sheet->getCell('J'.$y)->getValue(); 
                          $mprice = $sheet->getCell('K'.$y)->getValue(); 
                          $mtotal = $sheet->getCell('L'.$y)->getCalculatedValue(); 
                          $memo   = 'e7';
                          $total  = $sheet->getCell('O'.$y)->getCalculatedValue();         
                      } else if ($c6=='DESCRIPTION') {                             
                          $rx     = $sheet->getCell('A'.$y)->getValue();     
                          $pname  = $sheet->getCell('C'.$y)->getValue();     
                          $pqty   = $sheet->getCell('D'.$y)->getValue();   
                          $pprice = $sheet->getCell('E'.$y)->getValue(); 
                          $ptotal = $sheet->getCell('F'.$y)->getCalculatedValue(); 
                          $mname  = $sheet->getCell('G'.$y)->getValue(); 
                          $mqty   = $sheet->getCell('H'.$y)->getValue(); 
                          $mprice = $sheet->getCell('I'.$y)->getValue(); 
                          $mtotal = $sheet->getCell('J'.$y)->getCalculatedValue();  
                          $memo   = 'c6';
                          $total  = $sheet->getCell('M'.$y)->getCalculatedValue();        
                      } else if ($d6=='DESCRIPTION') {     
                          $rx     = $sheet->getCell('A'.$y)->getValue();     
                          $pname  = $sheet->getCell('D'.$y)->getValue();     
                          $pqty   = $sheet->getCell('E'.$y)->getValue();   
                          $pprice = $sheet->getCell('F'.$y)->getValue(); 
                          $ptotal = $sheet->getCell('G'.$y)->getCalculatedValue(); 
                          $mname  = $sheet->getCell('H'.$y)->getValue(); 
                          $mqty   = $sheet->getCell('I'.$y)->getValue(); 
                          $mprice = $sheet->getCell('J'.$y)->getValue(); 
                          $mtotal = $sheet->getCell('K'.$y)->getCalculatedValue(); 
                          $memo   = 'd6';
                          $total  = $sheet->getCell('N'.$y)->getCalculatedValue();            
                      } else if ($e6=='DESCRIPTION') {     
                          $rx     = $sheet->getCell('A'.$y)->getValue();     
                          $pname  = $sheet->getCell('E'.$y)->getValue();     
                          $pqty   = $sheet->getCell('F'.$y)->getValue();   
                          $pprice = $sheet->getCell('G'.$y)->getValue(); 
                          $ptotal = $sheet->getCell('H'.$y)->getCalculatedValue();  
                          $mname  = $sheet->getCell('I'.$y)->getValue(); 
                          $mqty   = $sheet->getCell('J'.$y)->getValue(); 
                          $mprice = $sheet->getCell('K'.$y)->getValue(); 
                          $mtotal = $sheet->getCell('L'.$y)->getCalculatedValue(); 
                          $memo   = 'e6';
                          $total  = $sheet->getCell('O'.$y)->getCalculatedValue();             
                      } else if ($f6=='DESCRIPTION') {     
                          if ($q6=='TOTAL'){
                              $rx     = $sheet->getCell('A'.$y)->getValue();     
                              $pname  = $sheet->getCell('F'.$y)->getValue();     
                              $pqty   = $sheet->getCell('G'.$y)->getValue();   
                              $pprice = $sheet->getCell('H'.$y)->getValue(); 
                              $ptotal = $sheet->getCell('I'.$y)->getCalculatedValue(); 
                              $mname  = $sheet->getCell('J'.$y)->getValue(); 
                              $mqty   = $sheet->getCell('K'.$y)->getValue(); 
                              $mprice = $sheet->getCell('L'.$y)->getValue(); 
                              $mtotal = $sheet->getCell('M'.$y)->getCalculatedValue(); 
                              $memo   = 'f6q6';
                              $total  = $sheet->getCell('Q'.$y)->getCalculatedValue();          
                          } else {
                              $rx     = $sheet->getCell('A'.$y)->getValue();     
                              $pname  = $sheet->getCell('F'.$y)->getValue();     
                              $pqty   = $sheet->getCell('G'.$y)->getValue();   
                              $pprice = $sheet->getCell('H'.$y)->getValue(); 
                              $ptotal = $sheet->getCell('I'.$y)->getCalculatedValue(); 
                              $mname  = $sheet->getCell('J'.$y)->getValue(); 
                              $mqty   = $sheet->getCell('K'.$y)->getValue(); 
                              $mprice = $sheet->getCell('L'.$y)->getValue(); 
                              $mtotal = $sheet->getCell('M'.$y)->getCalculatedValue();  
                              $memo   = 'f6p6';
                              $total  = $sheet->getCell('P'.$y)->getCalculatedValue();                
                          }
                      } else if ($g6=='DESCRIPTION') {     
                          if ($t6=='TOTAL'){    
                              $rx     = $sheet->getCell('A'.$y)->getValue();     
                              $pname  = $sheet->getCell('G'.$y)->getValue();     
                              $pqty   = $sheet->getCell('H'.$y)->getValue();   
                              $pprice = $sheet->getCell('I'.$y)->getValue(); 
                              $ptotal = $sheet->getCell('J'.$y)->getCalculatedValue(); 
                              $mname  = $sheet->getCell('K'.$y)->getValue(); 
                              $mqty   = $sheet->getCell('L'.$y)->getValue(); 
                              $mprice = $sheet->getCell('M'.$y)->getValue(); 
                              $mtotal = $sheet->getCell('N'.$y)->getCalculatedValue(); 
                              $memo   = 'g7t6';
                              $total  = $sheet->getCell('T'.$y)->getCalculatedValue();   
                          } else {    
                              $rx     = $sheet->getCell('A'.$y)->getValue();     
                              $pname  = $sheet->getCell('G'.$y)->getValue();     
                              $pqty   = $sheet->getCell('H'.$y)->getValue();   
                              $pprice = $sheet->getCell('I'.$y)->getValue(); 
                              $ptotal = $sheet->getCell('J'.$y)->getCalculatedValue(); 
                              $mname  = $sheet->getCell('K'.$y)->getValue(); 
                              $mqty   = $sheet->getCell('L'.$y)->getValue(); 
                              $mprice = $sheet->getCell('M'.$y)->getValue(); 
                              $mtotal = $sheet->getCell('N'.$y)->getCalculatedValue();  
                              $memo   = 'g7q6';
                              $total  = $sheet->getCell('Q'.$y)->getCalculatedValue();  
                          }                      
                      } else {
                          $msg.="$filename UnknowFormat!!";
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
                  $queryaa = "insert into log values ('$filename',$iii)"; 
                  $resultaa = mysql_query($queryaa);  
                  commit;
                  
                  echo ':'.$iii . "   ";                 
                  unset($PHPReader);
              }
            }
        } 
    }  
} 
$msg.='Done'; 
echo $msg;


?>
