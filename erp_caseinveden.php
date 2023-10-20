<?
  session_start();
  $pagtitle = "廠務部 &raquo; 期間內未出貨工單"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_caseindailytrace.php");
  if (is_null($_GET['bdate'])) {
    $bdate = date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }  
  
  if (is_null($_GET['edate'])) {
    $edate = date('Y-m-d');
  } else {
    $edate=$_GET['edate'];
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
                ->setCellValue('A1', $thisyear.' 期間內未出貨工單');       
                
    $y=1;                    
    $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.$y ,  '序號') 
                      ->setCellValue('B'.$y ,  '製處')           
                      ->setCellValue('C'.$y ,  '到貨日期')             
                      ->setCellValue('D'.$y ,  '客戶') 
                      ->setCellValue('E'.$y ,  'RX #') 
                      ->setCellValue('F'.$y ,  '工單號碼') 
                      ->setCellValue('G'.$y ,  '訂單號碼')  
                      ->setCellValue('H'.$y ,  '品代')  
                      ->setCellValue('I'.$y ,  '品名')   
                      ->setCellValue('J'.$y ,  '顆數') 
                      ->setCellValue('K'.$y ,  '傳真扣留期間') ;  
    $s1= "select occ01, occ02, sfb01, sfb08, sfb22, sfb221, sfb82, gem02, sfbud02, sfb05, ima02, oea02, oea14, gen02, to_char(tc_ohf004,'mm-dd') tc_ohf004, to_char(tc_ohf008,'mm-dd') tc_ohf008 from " .
           "(select occ01, occ02, sfb01, sfb08, sfb22, sfb221, sfb82, gem02, sfbud02, sfb05, ima02, to_char(oea02,'yy-mm-dd') oea02, oea14, gen02 from sfb_file, ima_file, oea_file, occ_file, gen_file, gem_file " .
           "where sfb28 is null and sfb04<7 and sfb22=oea01 and oea04=occ01 and oea14=gen01 and oea02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd')  " .
           "and sfb05=ima01 and ta_ima003!='Y'   and ta_ima004!='Y'  and ta_ima005!='Y' " .
           "and sfb01 not in ( select sfl02 from sfl_file) "  .          
           "and sfb01 not in ( select tc_ogb001 from tc_ogb_file) " .
           "and sfb82=gem01 ) ".
           "left join tc_ohf_file on sfb01=tc_ohf001 " . 
           "order by sfb82, oea02, occ01, sfbud02 " ;    
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);  
    $y=1;  
    $i=1;
    $oldgem02='';                                
    while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {         
          if ($oldgem02 != $row1['GEM02'] ) {
              $y++;
              $i=1;
              $oldgem02=$row1['GEM02'];
          }
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.$y ,  $i)   
                      ->setCellValue('B'.$y ,  $row1['GEM02']) 
                      ->setCellValue('C'.$y ,  $row1['OEA02']) 
                      ->setCellValue('D'.$y ,  $row1['OCC01'] .' ' .$row1['OCC02']) 
                      ->setCellValue('E'.$y ,  "' ".$row1['SFBUD02']) 
                      ->setCellValue('F'.$y ,  $row1['SFB01']) 
                      ->setCellValue('G'.$y ,  $row1['SFB22']. ' ' . $row1['SFB221'])
                      ->setCellValue('H'.$y ,  $row1['SFB05'])
                      ->setCellValue('I'.$y ,  $row1['IMA02'])
                      ->setCellValue('J'.$y ,  $row1['SFB08']) 
                      ->setCellValue('K'.$y ,  $row1['TC_OHF004'].'  '.$row1['TC_OHF008']) ;
          $y++;    
          $i++;                                                                        
    }                      
                
    $objPHPExcel->getActiveSheet()->setTitle('期間內未出貨工單');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $bdate. '_'.$edate.'_caseinveden.xls"');   
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
<link href="css.css" rel="stylesheet" type="text/css">
<p>期間內未出貨工單 </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            到貨期間:   
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$bdate;?>> ~~                                            
            <input name="edate" type="text" id="edate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$edate;?>>&nbsp; &nbsp;
            <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;     
            <input type="submit" name="submit" id="submit" value="匯出">     
            </div></td>        
        </tr>
    </table>
  </div>
</form>

<? if (is_null($_GET['submit'])) die ; ?>                      
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>製處</th>  
    <th>Order Date</th>  
    <th>客戶</th>          
    <th>RX #</th>     
    <th>工單號碼</th>   
    <th>訂單號碼</th>    
    <th>品代</th>   
    <th>品名</th>                                                                                                    
    <th>顆數</th>    
    <th>傳真扣留期間</th> 
  </tr>
  <?          
      $s1= "select occ01, occ02, sfb01, sfb08, sfb22, sfb221, sfb82, gem02, sfbud02, sfb05, ima02, oea02, oea14, gen02, to_char(tc_ohf004,'mm-dd') tc_ohf004, to_char(tc_ohf008,'mm-dd') tc_ohf008 from " .
           "(select occ01, occ02, sfb01, sfb08, sfb22, sfb221, sfb82, gem02, sfbud02, sfb05, ima02, to_char(oea02,'yy-mm-dd') oea02, oea14, gen02 from sfb_file, ima_file, oea_file, occ_file, gen_file, gem_file " .
           "where sfb28 is null and sfb04<7 and sfb22=oea01 and oea04=occ01 and oea14=gen01 and oea02 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd')  " .
           "and sfb05=ima01 and ta_ima003!='Y'   and ta_ima004!='Y'  and ta_ima005!='Y' " .
           "and sfb01 not in ( select sfl02 from sfl_file) "  .          
           "and sfb01 not in ( select tc_ogb001 from tc_ogb_file) " .
           "and sfb82=gem01 ) ".
           "left join tc_ohf_file on sfb01=tc_ohf001 " . 
           "order by sfb82, oea02, occ01, sfbud02 " ;    
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);   
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
      ?>    
	      <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
          <td><?=$row1["GEM02"];?></td>  
          <td><?=$row1["OEA02"];?></td> 
          <td><?=$row1["OCC01"].' '.$row1["OCC02"];?></td>
          <td><?=$row1["SFBUD02"];?></td> 
          <td><?=$row1["SFB01"];?></td>  
          <td><?=$row1["SFB22"];?> <?=$row1["SFB221"];?> </td>  
          <td><?=$row1["SFB05"];?></td>  
          <td><?=$row1["IMA02"];?></td>   
          <td><?=$row1["SFB08"];?></td>  
          <td><?=$row1["TC_OHF004"];?> <?=$row1["TC_OHF008"];?> </td>   
    <?  
      }   
  ?>            
</table>  
</form>
