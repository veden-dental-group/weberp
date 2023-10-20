<?
  session_start();
  $pagtitle = "廠務部 &raquo; 每日到貨CASE刷卡報工出貨追蹤--有到貨未刷進"; 
  include("_data.php");
  include("_erp.php");
  // auth("erp_caseindailytrace1.php");   
  $tdate=$_GET['tdate'];      
  $maker=$_GET['maker']; 
  //$issent=$_GET['issent']; 
  $issent='';
  $ecb06=$_GET['ecb06'];   
  $isecb06=$_GET['isecb06']; 
  
  if ($issent=='Y') {
    $issentfilter='';  
  } else {
    $issentfilter = ' and ogb13 is null ';  
  }    
  
  if ($isecb06=='Y') {
    $isecb06filter='';  
  } else {
    $isecb06filter = " and ecb06='$ecb06' ";  
  }  
  
  if ($maker==''){
     $makerfilter='';    
  } else {
     $makerfilter=" and sfb82='$maker' ";  
  }

  if ($_GET["submit"]=="匯出") {   

  //$filename='templates/casesoutmonthlysummary_byarea.xls';
      
  error_reporting(E_NONE);  
  require_once 'classes/PHPExcel.php'; 
  require_once 'classes/PHPExcel/IOFactory.php';  
  $objReader = PHPExcel_IOFactory::createReader('Excel5');
  //$objPHPExcel = $objReader->load($filename);  
  $objPHPExcel = new PHPExcel(); 
  // Set properties
  $objPHPExcel ->getProperties()->setCreator('Frank' )
               ->setLastModifiedBy('Frank')
               ->setTitle('Frank')
               ->setSubject('Frank')
               ->setDescription('Frank')
               ->setKeywords('Frank')
               ->setCategory('Frank');  
               
  $objPHPExcel->setActiveSheetIndex(0)
              ->setCellValue('A1', $thisyear.' 每日到貨CASE刷卡報工出貨追蹤--有到貨未刷進');       
              
  $y=1;                    
  $objPHPExcel->setActiveSheetIndex(0) 
                    ->setCellValue('A'.$y ,  '到貨日期')   
                    ->setCellValue('B'.$y ,  'RX #')  
                    ->setCellValue('C'.$y ,  '工單號碼') 
                    ->setCellValue('D'.$y ,  '訂單號碼')                    
                    ->setCellValue('E'.$y ,  '製處') 
                    ->setCellValue('F'.$y ,  '品代') 
                    ->setCellValue('G'.$y ,  '品名') 
                    ->setCellValue('H'.$y ,  '工序') 
                    ->setCellValue('I'.$y ,  '在製量') 
                    ->setCellValue('J'.$y ,  '刷進日期')
                    ->setCellValue('K'.$y ,  '刷進時間') 
                    ->setCellValue('L'.$y ,  '刷出日期') 
                    ->setCellValue('M'.$y ,  '刷出時間') 
                    ->setCellValue('N'.$y ,  '傳真日期')  
                    ->setCellValue('O'.$y ,  '是否出貨') ;
                      
  $s1="select oea02, sfbud02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, ima02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011, ogb03, tc_ohf004 from " .  
              "(select to_char(oea02,'yy/mm/dd') oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02, ecb06, ecb17, tc_srg004, tc_srg005, to_char(tc_srg007,'yy/mm/dd') tc_srg007, tc_srg008, to_char(tc_srg010,'yy/mm/dd') tc_srg010, tc_srg011, ogb03 from " .
                  "(select  oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                      "(select  oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                          "(select oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02 " .
                          "from sfb_file, gem_file, ima_file, oea_file " .
                          "where sfb22=oea01 $makerfilter and oea02=to_date('$tdate','yy/mm/dd') " .
                          "and sfb82=gem01 and sfb05=ima01) " .   
                  "left join tc_srg_file  on  sfb01=tc_srg001 ) , ecb_file where sfb05=ecb01 and tc_srg004=ecb03 ) " .
              "left join ogb_file on sfb22=ogb31 and sfb221=ogb32 where tc_srg007  not null $issentfilter $isecb06filter ) " .
          "left join ( select distinct tc_ohf001, concat(to_char(tc_ohf004,'yy/mm/dd'),gen02) tc_ohf004 from tc_ohf_file, gen_file where tc_ohf006=gen01 ) on sfb01=tc_ohf001 order by oea02, sfbud02, sfb01, sfb22, sfb221";     
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);  
    $y=2;                                  
    while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
          if (is_null($row1['OGB03'])) {
            $ogb03='';
        } else {
            $ogb03="已出";
        }           
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.$y ,  $row1['OEA02'])
                    ->setCellValue('B'.$y ,  $row1['SFBUD02'])  
                    ->setCellValue('C'.$y ,  $row1['SFB01']) 
                    ->setCellValue('D'.$y ,  $row1['SFB22']. ' ' .$row1['SFB221'])                     
                    ->setCellValue('E'.$y ,  $row1['GEM02']) 
                    ->setCellValue('F'.$y ,  $row1['SFB05']) 
                    ->setCellValue('G'.$y ,  $row1['IMA02']) 
                    ->setCellValue('H'.$y ,  $row1['ECB17']) 
                    ->setCellValue('I'.$y ,  $row1['TC_SRG005'])                      
                    ->setCellValue('J'.$y ,  $row1['TC_SRG007']) 
                    ->setCellValue('K'.$y ,  $row1['TC_SRG008']) 
                    ->setCellValue('L'.$y ,  $row1['TC_SRG010']) 
                    ->setCellValue('M'.$y ,  $row1['TC_SRG011']) 
                    ->setCellValue('N'.$y ,  $row1['TC_OHF004'])  
                    ->setCellValue('O'.$y ,  $ogb03) ;
          $y++;                                                                            
    }                      
                
    $objPHPExcel->getActiveSheet()->setTitle('每日到貨CASE刷卡報工出貨追蹤--有刷進未刷出');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $maker. '_'.$tdate.'_casetrace.xls"');   
    header('Cache-Control: max-age=0'); 
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
    $objWriter->save('php://output'); 
    exit;    
  }   
  
  include("_header.php");
?>
<script language='JavaScript'>
checked = false;
function checkedAll () {
  if (checked == false) {
    checked = true
  }else{
    checked = false
  }
  
  for (var i = 0; i < document.form1.elements.length; i++) {
    var e = document.form1.elements[i];
        if (e.type == 'checkbox' && e.disabled==false) {
            e.checked = checked
        }                                                          
  }  
}
</script>
                    
<form action="<?=$PHP_SELF;?>" method="get" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>到貨日期</th> 
    <th>RX #</th>    
    <th>工單號碼</th>     
    <th>訂單號碼</th>  
    <th>製處</th> 
    <th>品代</th>   
    <th>品名</th>   
    <th>工序</th>  
    <th>在製量</th>
    <th>刷進日期</th>   
    <th>刷進時間</th> 
    <th>刷出日期</th> 
    <th>刷出時間</th>   
    <th>傳真日期</th>                                                                                        
    <th>是否出貨</th> 
  </tr>
  <?          
      $s1="select oea02, sfbud02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, ima02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011, ogb03, tc_ohf004 from " .  
              "(select to_char(oea02,'yy/mm/dd') oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02, ecb06, ecb17, tc_srg004, tc_srg005, to_char(tc_srg007,'yy/mm/dd') tc_srg007, tc_srg008, to_char(tc_srg010,'yy/mm/dd') tc_srg010, tc_srg011, ogb03 from " .
                  "(select  oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                      "(select  oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                          "(select oea02, sfb01, sfb22, sfb221, sfb82, gem02, sfb05, sfbud02, ima02 " .
                          "from sfb_file, gem_file, ima_file, oea_file " .
                          "where sfb22=oea01 $makerfilter and oea02=to_date('$tdate','yy/mm/dd') " .
                          "and sfb82=gem01 and sfb05=ima01) " .   
                  "left join tc_srg_file  on  sfb01=tc_srg001 ) , ecb_file where sfb05=ecb01 and tc_srg004=ecb03 ) " .
              "left join ogb_file on sfb22=ogb31 and sfb221=ogb32 where tc_srg007 is null $issentfilter $isecb06filter ) " .
          "left join ( select distinct tc_ohf001, concat(to_char(tc_ohf004,'yy/mm/dd'),gen02) tc_ohf004 from tc_ohf_file, gen_file where tc_ohf006=gen01 ) on sfb01=tc_ohf001 order by oea02, sfbud02, sfb01, sfb22, sfb221";
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);   
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
        if (is_null($row1['OGB03'])) {
            $ogb03='';
        } else {
            $ogb03="已出";
        }  
        if (is_null($row1['TC_OHF004'])) {
            $tc_ohf004='';
        } else {
            $tc_ohf004="已出";
        } 
  ?>    
	      <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">            
          <td><?=$row1["OEA02"];?></td> 
          <td><?=$row1["SFBUD02"];?></td>  
          <td><?=$row1["SFB01"];?></td>  
          <td><?=$row1["SFB22"].'-'.$row1['SFB221'];?></td>            
          <td><?=$row1["GEM02"];?></td> 
          <td><?=$row1["SFB05"];?></td>
          <td><?=$row1["IMA02"];?></td>
          <td><?=$row1["ECB06"].' '.$row1['ECB17'];?></td>
          <td><?=$row1["TC_SRG005"];?></td>       
          <td><?=$row1["TC_SRG007"];?></td>
          <td><?=$row1["TC_SRG008"];?></td>   
          <td><?=$row1["TC_SRG010"];?></td>   
          <td><?=$row1["TC_SRG011"];?></td>   
          <td><?=$row1["TC_OHF004"];?></td> 
          <td><?=$ogb03;?></td>   
        </tr>
  <?         
      }   
  ?>              
</table>  
<input type="hidden" name="tdate" id="tdate" value="<?=$tdate;?>">
<input type="hidden" name="maker" id="maker" value="<?=$maker;?>"> 
<input type="hidden" name="issent" id="issent" value="<?=$issent;?>"> 
<input type="hidden" name="isecb06" id="isecb06" value="<?=$isecb06;?>"> 
<input type="hidden" name="ecb06" id="ecb06" value="<?=$ecb06;?>"> 
<input type="submit" name="submit" id="submit" value="匯出">  
</form>
