<?php
  session_start();
  $pagetitle = "財務部 &raquo; 各年到貨量分析";
  include("_data.php");
  auth("erp_casesinyearlysummary.php");  
  
  if (is_null($_GET['thisyear'])) {
    $thisyear = date("Y") ;     
  } else {      
    $thisyear = $_GET['thisyear'] ;   
  }  
  
  if ($_GET["submit"]=="匯出") {   

    $filename='templates/casesinyearlysummary.xls';
        
    error_reporting(E_NONE);  
    require_once 'classes/PHPExcel.php'; 
    require_once 'classes/PHPExcel/IOFactory.php';  
    $objReader = PHPExcel_IOFactory::createReader('Excel5');
    $objPHPExcel = $objReader->load($filename);  
    // Set properties
    $objPHPExcel ->getProperties()->setCreator('Frank' )
                 ->setLastModifiedBy('Frank')
                 ->setTitle('Frank')
                 ->setSubject('Frank')
                 ->setDescription('Frank')
                 ->setKeywords('Frank')
                 ->setCategory('Frank');  
                 
    $objPHPExcel->setActiveSheetIndex(0)
                ->setCellValue('A1', '各年到货量分析');       
                      
    $y=3;      
    $s2008="select '2008' oea02, 30546 m1, 35487 m2, 30576 m3, 34728 m4, 39464 m5, 42030 m6, 40726 m7, 41310 m8, 40402 m9, 44315 m10, 44567 m11, 37038 m12 from dual "; 
    $s2009="select '2009' oea02, 46482 m1, 49445 m2, 55065 m3, 53629 m4, 53287 m5, 51238 m6, 46434 m7, 38504 m8, 55196 m9, 50205 m10, 46948 m11, 27261 m12 from dual "; 
    $s2010="select '2010' oea02, 39734 m1, 45157 m2, 44300 m3, 42728 m4, 45001 m5, 44813 m6, 43047 m7, 36955 m8, 42804 m9, 50551 m10, 47555 m11, 42641 m12 from dual "; 
    $s2011="select '2011' oea02, 51049 m1, 47800 m2, 55908 m3, 55011 m4, 62615 m5, 58524 m6, 54882 m7, 50452 m8, 55396 m9, 54615 m10, 52334 m11, 51269 m12 from dual "; 
    $s2012="select substr(oea02,1,4) oea02 , decode(substr(oea02,6,2),'01', total, 0) m1,  decode(substr(oea02,6,2),'02', total, 0) m2,  decode(substr(oea02,6,2),'03', total, 0) m3, ".   
                                     "decode(substr(oea02,6,2),'04', total, 0) m4,  decode(substr(oea02,6,2),'05', total, 0) m5,  decode(substr(oea02,6,2),'06', total, 0) m6, ". 
                                     "decode(substr(oea02,6,2),'07', total, 0) m7,  decode(substr(oea02,6,2),'08', total, 0) m8,  decode(substr(oea02,6,2),'09', total, 0) m9, ". 
                                     "decode(substr(oea02,6,2),'10', total, 0) m10, decode(substr(oea02,6,2),'11', total, 0) m11, decode(substr(oea02,6,2),'12', total, 0) m12 from " .   
          " (select oea02, sum(total) total from (select to_char(oea02, 'yyyy-mm') oea02, (oeb12 * imaud07) total from oeb_file, oea_file, ima_file where oeb01=oea01 and oeb04=ima01 and to_char(oea02,'yyyy')>'2011' and oeaconf='Y' ) group by oea02 ) ";
    
    $s2   ="select oea02, sum(m1) m1, sum(m2) m2, sum(m3) m3,  sum(m4) m4, sum(m5) m5, sum(m6) m6, sum(m7) m7, sum(m8) m8, sum(m9) m9, sum(m10) m10, sum(m11) m11, sum(m12) m12 from  " .  
           "( $s2008 union all $s2009 union all $s2010 union all $s2011 union all $s2012 ) " .
           "group by oea02 order by oea02 ";   
    $erp_sql2 = oci_parse($erp_conn,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'. $y, $row2['OEA02'].'年')
                      ->setCellValue('B'. $y, $row2["M1"])
                      ->setCellValue('C'. $y, $row2["M2"])
                      ->setCellValue('D'. $y, $row2["M3"])   
                      ->setCellValue('E'. $y, $row2["M4"])
                      ->setCellValue('F'. $y, $row2["M5"])
                      ->setCellValue('G'. $y, $row2["M6"])
                      ->setCellValue('H'. $y, $row2["M7"])
                      ->setCellValue('I'. $y, $row2["M8"])
                      ->setCellValue('J'. $y, $row2["M9"])
                      ->setCellValue('K'. $y, $row2["M10"])
                      ->setCellValue('L'. $y, $row2["M11"])
                      ->setCellValue('M'. $y, $row2["M12"]);
          $y++;                                                   
    }                      
                
    //$objPHPExcel->getActiveSheet()->setTitle('各年到货量分析');
    
    //導出第二個Sheet中 Veden-Group的資料                         
    $y=3;                                                                                                                                                             
    $s2011pls="select 'PLS'  name , '2011' oea02, 13086 m1, 11341 m2, 11512 m3, 12356 m4, 15854 m5, 13297 m6, 14221 m7, 12876 m8, 12791 m9, 13747 m10, 12633 m11, 14803 m12 from dual "; 
    $s2011fdl="select 'FDL'  name , '2011' oea02, 4050  m1,  3528 m2,  4407 m3,  4602 m4,  4747 m5,  4616 m6,  4699 m7,  4496 m8,  4389 m9,  4593 m10,  4143 m11,  4723 m12 from dual "; 
    $s2011ed ="select 'ED'   name , '2011' oea02, 392   m1,   424 m2,   389 m3,   470 m4,   674 m5,   625 m6,   666 m7,   586 m8,   495 m9,   636 m10,   452 m11,   727 m12 from dual "; 
    $s2011ids="select 'IDS'  name , '2011' oea02, 251   m1,   262 m2,   228 m3,   229 m4,   317 m5,   345 m6,   422 m7,   530 m8,   531 m9,   546 m10,   353 m11,   393 m12 from dual "; 
    $s2011vla="select 'VL-A' name , '2011' oea02, 0     m1,     0 m2,     0 m3,   247 m4,   364 m5,   527 m6,   650 m7,   761 m8,   918 m9,   741 m10,  1138 m11,   826 m12 from dual "; 
    $s2012="select decode(oea04,'U132','PLS','U115','FDL','U134','ED','U142','IDS','A114','VL-A','') name, substr(oea02,1,4) oea02 , ".
                                     "decode(substr(oea02,6,2),'01', total, 0) m1,  decode(substr(oea02,6,2),'02', total, 0) m2,  decode(substr(oea02,6,2),'03', total, 0) m3, ".   
                                     "decode(substr(oea02,6,2),'04', total, 0) m4,  decode(substr(oea02,6,2),'05', total, 0) m5,  decode(substr(oea02,6,2),'06', total, 0) m6, ". 
                                     "decode(substr(oea02,6,2),'07', total, 0) m7,  decode(substr(oea02,6,2),'08', total, 0) m8,  decode(substr(oea02,6,2),'09', total, 0) m9, ". 
                                     "decode(substr(oea02,6,2),'10', total, 0) m10, decode(substr(oea02,6,2),'11', total, 0) m11, decode(substr(oea02,6,2),'12', total, 0) m12 from " .   
          " (select oea04, oea02, sum(total) total from (select substr(oea04,1,4) oea04, to_char(oea02, 'yyyy-mm') oea02, (oeb12 * imaud07) total from oeb_file, oea_file, ima_file where oeb01=oea01 and oeb04=ima01 and to_char(oea02,'yyyy')>'2011' and oeaconf='Y' and substr(oea04,1,4) in ('U132','U115','U134','U142','A114') ) group by oea04, oea02 ) ";
    
    $s2   ="select name, oea02, sum(m1) m1, sum(m2) m2, sum(m3) m3,  sum(m4) m4, sum(m5) m5, sum(m6) m6, sum(m7) m7, sum(m8) m8, sum(m9) m9, sum(m10) m10, sum(m11) m11, sum(m12) m12 from  " .  
           "( $s2011pls union all $s2011fdl union all $s2011ed union all $s2011ids union all $s2011vla union all $s2012 ) " .
           "group by name, oea02 order by name, oea02 ";   
    $erp_sql2 = oci_parse($erp_conn1,$s2 );
    oci_execute($erp_sql2);   
    while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $objPHPExcel->setActiveSheetIndex(1)
                      ->setCellValue('A'. $y, $row2['NAME']) 
                      ->setCellValue('B'. $y, $row2['OEA02'])
                      ->setCellValue('C'. $y, $row2["M1"])
                      ->setCellValue('D'. $y, $row2["M2"])
                      ->setCellValue('E'. $y, $row2["M3"])   
                      ->setCellValue('F'. $y, $row2["M4"])
                      ->setCellValue('G'. $y, $row2["M5"])
                      ->setCellValue('H'. $y, $row2["M6"])
                      ->setCellValue('I'. $y, $row2["M7"])
                      ->setCellValue('J'. $y, $row2["M8"])
                      ->setCellValue('K'. $y, $row2["M9"])
                      ->setCellValue('L'. $y, $row2["M10"])
                      ->setCellValue('M'. $y, $row2["M11"])
                      ->setCellValue('N'. $y, $row2["M12"]);
          $y++;                                                   
    }                      
                
    //$objPHPExcel->getActiveSheet()->setTitle('Veden-Group'); 
    
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $thisyear . '_yearlysummary.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下為各年到貨量分析表 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        年份:   
        <input name="thisyear" type="text" id="thisyear" onfocus="WdatePicker({dateFmt:'yyyy'})" value="<?=$thisyear;?>"> &nbsp;&nbsp;  
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;     
        <input type="submit" name="submit" id="submit" value="匯出">         
      </td></tr>
    </table>
  </div>
</form>  
          
<? if (is_null($_GET['submit'])) die ; ?>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th>
        <th>年份</th>   
        <th style="text-align:right">一月</th>  
        <th style="text-align:right">二月</th> 
        <th style="text-align:right">三月</th>
        <th style="text-align:right">四月</th>
        <th style="text-align:right">五月</th>  
        <th style="text-align:right">六月</th>    
        <th style="text-align:right">七月</th>        
        <th style="text-align:right">八月</th>    
        <th style="text-align:right">九月</th> 
        <th style="text-align:right">十月</th>    
        <th style="text-align:right">十一月</th>  
        <th style="text-align:right">十二月</th>  
        <th style="text-align:right">合計</th>
    </tr>
    <?
      //算出要秀出的年份 置於陣列中
      $yeara=array();
      $x=1;
      for ($y=2012;$y<=$thisyear; $y++){
          $yeara[$x]=$y; 
          $x++;
      }
      $s2008="select '2008' oea02, 30546 m1, 35487 m2, 30576 m3, 34728 m4, 39464 m5, 42030 m6, 40726 m7, 41310 m8, 40402 m9, 44315 m10, 44567 m11, 37038 m12 from dual "; 
      $s2009="select '2009' oea02, 46482 m1, 49445 m2, 55065 m3, 53629 m4, 53287 m5, 51238 m6, 46434 m7, 38504 m8, 55196 m9, 50205 m10, 46948 m11, 27261 m12 from dual "; 
      $s2010="select '2010' oea02, 39734 m1, 45157 m2, 44300 m3, 42728 m4, 45001 m5, 44813 m6, 43047 m7, 36955 m8, 42804 m9, 50551 m10, 47555 m11, 42641 m12 from dual "; 
      $s2011="select '2011' oea02, 51049 m1, 47800 m2, 55908 m3, 55011 m4, 62615 m5, 58524 m6, 54882 m7, 50452 m8, 55396 m9, 54615 m10, 52334 m11, 51269 m12 from dual "; 
      $s2012="select substr(oea02,1,4) oea02 , decode(substr(oea02,6,2),'01', total, 0) m1,  decode(substr(oea02,6,2),'02', total, 0) m2,  decode(substr(oea02,6,2),'03', total, 0) m3, ".   
                                       "decode(substr(oea02,6,2),'04', total, 0) m4,  decode(substr(oea02,6,2),'05', total, 0) m5,  decode(substr(oea02,6,2),'06', total, 0) m6, ". 
                                       "decode(substr(oea02,6,2),'07', total, 0) m7,  decode(substr(oea02,6,2),'08', total, 0) m8,  decode(substr(oea02,6,2),'09', total, 0) m9, ". 
                                       "decode(substr(oea02,6,2),'10', total, 0) m10, decode(substr(oea02,6,2),'11', total, 0) m11, decode(substr(oea02,6,2),'12', total, 0) m12 from " .   
             " (select oea02, sum(total) total from (select to_char(oea02, 'yyyy-mm') oea02, (oeb12 * imaud07) total from oeb_file, oea_file, ima_file where oeb01=oea01 and oeb04=ima01 and to_char(oea02,'yyyy')>'2011' and oeaconf='Y' ) group by oea02 ) ";
      
      $s2   ="select oea02, sum(m1) m1, sum(m2) m2, sum(m3) m3,  sum(m4) m4, sum(m5) m5, sum(m6) m6, sum(m7) m7, sum(m8) m8, sum(m9) m9, sum(m10) m10, sum(m11) m11, sum(m12) m12 from  " .  
             "( $s2008 union all $s2009 union all $s2010 union all $s2011 union all $s2012 ) " .
             "group by oea02 order by oea02 ";                                                                                                                                   
        
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";             
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16">  
              <td><?=$row2["OEA02"];?></td>   
		          <td style="text-align:right"><?=number_format($row2["M1"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["M2"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M3"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M4"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M5"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M6"], 2);?></td>                                                                      
              <td style="text-align:right"><?=number_format($row2["M7"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M8"], 2);?></td> 
              <td style="text-align:right"><?=number_format($row2["M9"], 2);?></td>  
              <td style="text-align:right"><?=number_format($row2["M10"], 2);?></td>   
              <td style="text-align:right"><?=number_format($row2["M11"], 2);?></td>
              <td style="text-align:right"><?=number_format($row2["M12"], 2);?></td>                                                        
              <td style="text-align:right"><?=number_format($row2["M1"]+$row2["M2"]+$row2["M3"]+$row2["M4"]+$row2["M5"]+$row2["M6"]+$row2["M7"]+$row2["M8"]+$row2["M9"]+$row2["M10"]+$row2["M11"]+$row2["M12"] , 2);?></td> 
          </tr>
		  <?
			}
      ?>       
</table>    