<?php
  session_start();
  $pagetitle = "業務部 &raquo; Delay Report";
  include("_data.php");
  include("_erp.php");
  auth("erp_clientdelayreport.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d',strtotime('-3 days'));
  } else {
    $bdate=$_GET['bdate'];
  }                              
  
  if (is_null($_GET['occ01'])) {
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }  
  
    if ($_GET["submit"]=="匯出") {   
        error_reporting(E_ALL);  
        require_once 'classes/PHPExcel.php'; 
        require_once 'classes/PHPExcel/IOFactory.php';  
        $objPHPExcel = new PHPExcel();
        // Set properties
        $objPHPExcel ->getProperties()->setCreator( 'DelayReport')
                     ->setLastModifiedBy('DelayReport')
                     ->setTitle('DelayReport')
                     ->setSubject('DelayReport')
                     ->setDescription('DelayReport')
                     ->setKeywords('DelayReport')
                     ->setCategory('DelayReport');
        
        // Add some data      
                        
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A1', 'Delay Report');                    
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A3', 'No.')  
                    ->setCellValue('B3', 'Received Date')                  
                    ->setCellValue('C3', 'Case No.')
                    ->setCellValue('D3', 'Order No.')                         
                    ->setCellValue('E3', 'Due Date')   
                    ->setCellValue('F3', 'Product Code')  
                    ->setCellValue('G3', 'Product Description')  
                    ->setCellValue('H3', 'Qty.')  
                    ->setCellValue('I3', 'Fax Date') 
                    ->setCellValue('J3', 'Response Date') 
                    ->setCellValue('K3', 'Comment');  

        $y=4; 
        $qtytotal=0;  
        $i=0;   
        $oldrxno='';     
        $oldoea01=''; 
        $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
        $s2= "select rxno, indate, duedate, oea04, oea01, oea02, sfb01, sfb08, gem02, ima01, ima02, ima1002 from " .
             "(select ta_oea006 rxno, to_char(oea02,'yyyy/mm/dd') indate, to_char(ta_oea005,'yyyy/mm/dd') duedate, oea04, oea01, oea02, sfb01, sfb08, gem02, ima01, ima02, ima1002  
               from oea_file, sfb_file, gem_file, ima_file where oea02 <= to_date('$bdate1','yymmdd') and oea04 in (select occ01 from occ_file where occ01='$occ01')  
               and oea01=sfb22 and sfb82=gem01 and sfb05=ima01 and ta_ima005='N' and sfb28 is null) a " . 
             "left join oga_file on oea01=oga16 where oga16 is null order by indate,rxno, oea01";
        $erp_sql2 = oci_parse($erp_conn1,$s2 );
        oci_execute($erp_sql2);  
        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {    
          //只要order no不一樣 就重新找一次工單 判斷有無fax
          $oea01=$row2['OEA01'];
          if ($oea01!=$oldoea01){    
              $faxout='';
              $faxin='';          
              $sfax="select to_char(tc_ohf004,'yyyy/mm/dd') tc_ohf004, to_char(tc_ohf008,'yyyy/mm/dd') tc_ohf008 from tc_ohf_file where tc_ohf001 in (select sfb01 from sfb_file where sfb22='$oea01') order by tc_ohf004 desc";
              $erp_sqlfax = oci_parse($erp_conn1,$sfax );
              oci_execute($erp_sqlfax); 
              $rowfax = oci_fetch_array($erp_sqlfax, OCI_ASSOC);
              if (!is_null($rowfax['TC_OHF004']))  $faxout=$rowfax['TC_OHF004'];  
              if (!is_null($rowfax['TC_OHF008']))  $faxin =$rowfax['TC_OHF008'];
              $oldoea01=$oea01;
          } 
          
          if ( $faxout!='' and $faxin=='') {
              //有faxout 沒有faxIn的就不要秀delay 因為是fax case
          } else {   
              $rxno=$row2['RXNO'];   
              $sfb01=$row2['SFB01'];   
              $gem02=$row2['GEM02'];
              $msg= findcasewithsfb01($sfb01,$gem02,$erp_conn1,$erp_conn2);            

              if ( $rxno!=$oldrxno) {
                  $i++;  
                  $ii=$i;
                  $rxno =$row2['RXNO'];
                  $oldrxno=$row2['RXNO'];
                  $indate=$row2['INDATE'];
                  $duedate=$row2['DUEDATE'];  
                } else {     
                  $rxno='';
                  $ii='';
                  $indate='';
                  $duedate='';
                }       
                $objPHPExcel->setActiveSheetIndex(0)
                            ->setCellValue('A'. $y, $ii)
                            ->setCellValue('B'. $y, $indate) 
                            ->setCellValue('C'. $y, $rxno)
                            ->setCellValue('D'. $y, $oea01)                                 
                            ->setCellValue('E'. $y, $duedate)   
                            ->setCellValue('F'. $y, $row2["IMA01"])   
                            ->setCellValue('G'. $y, $row2["IMA02"])   
                            ->setCellValue('H'. $y, $row2["SFB08"])   
                            ->setCellValue('I'. $y, $faxout)  
                            ->setCellValue('J'. $y, $faxin)  
                            ->setCellValue('K'. $y, $msg); 
                $y++;   
          } 
        }
        //total
        $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.($y+1), 'Total');  
                                                   
        $objPHPExcel->setActiveSheetIndex(0)        
                    ->setCellValue('H'.($y+1), '=sum(H3:H' . ($y-1). ")");                 
        
        // Rename sheet
        $objPHPExcel->getActiveSheet()->setTitle('DelayReport');
                                                                
        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');  
        header('Content-Disposition: attachment;filename="DelayReport.xls"');    
        header('Cache-Control: max-age=0'); 
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
        $objWriter->save('php://output'); 
        exit;    
  }  
    
    //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');
    
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>客戶Delay Report </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         到貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> 
        客戶: 
        <select name="occ01" id="occ01">  
            <?
              $s1= "select occ01,occ02 from occ_file where occ01 in ( select decode(occud03,'Y',occ07, occ09) from occ_file ) group by occ01, occ02 order by occ01 ";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC01"];  
                  if ($_GET["occ01"] == $row1["OCC01"]) echo " selected";                  
                  echo ">" . $row1['OCC01'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>   &nbsp;&nbsp;   &nbsp;&nbsp;  
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="匯出">  &nbsp;&nbsp;   &nbsp;&nbsp;            
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>Received Date</th>         
        <th>RX #</th>  
        <th>Order No.</th> 
        <th>Due Date</th> 
        <th>Product Code</th>
        <th>Product Description</th>   
        <th>Qty.</th>    
        <th>Fax Date</th>
        <th>Response Date</th>
        <th>Comment</th> 
    </tr>
    <?
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      //檢查工單號有無出貨  , 配件不用   
      $s2= "select rxno, indate, duedate, oea04, oea01, oea02, sfb01, sfb08, gem02, ima01, ima02, ima1002 from " .
           "(select ta_oea006 rxno, to_char(oea02,'yyyy/mm/dd') indate, to_char(ta_oea005,'yyyy/mm/dd') duedate, oea04, oea01, oea02, sfb01, sfb08, gem02, ima01, ima02, ima1002  
             from oea_file, sfb_file, gem_file, ima_file where oea02 <= to_date('$bdate1','yymmdd') and oea04 in (select occ01 from occ_file where occ01='$occ01')  
             and oea01=sfb22 and sfb82=gem01 and sfb05=ima01 and ta_ima005='N' and sfb28 is null) a " . 
           "left join oga_file on oea01=oga16 where oga16 is null order by indate,rxno, oea01";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;
      $oldrxno='';     
      $oldoea01='';
      $total=0;
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
          //只要order no不一樣 就重新找一次工單 判斷有無fax
          $oea01=$row2['OEA01'];
          if ($oea01!=$oldoea01){    
              $faxout='';
              $faxin='';          
              $sfax="select to_char(tc_ohf004,'yyyy/mm/dd') tc_ohf004, to_char(tc_ohf008,'yyyy/mm/dd') tc_ohf008 from tc_ohf_file where tc_ohf001 in (select sfb01 from sfb_file where sfb22='$oea01') order by tc_ohf004 desc";
              $erp_sqlfax = oci_parse($erp_conn1,$sfax );
              oci_execute($erp_sqlfax); 
              $rowfax = oci_fetch_array($erp_sqlfax, OCI_ASSOC);
              if (!is_null($rowfax['TC_OHF004']))  $faxout=$rowfax['TC_OHF004'];  
              if (!is_null($rowfax['TC_OHF008']))  $faxin =$rowfax['TC_OHF008'];
              $oldoea01=$oea01;
          } 
          
          if ( $faxout!='' and $faxin=='') {
              //有faxout 沒有faxIn的就不要秀delay 因為是fax case
          } else {
              $total+=$row2['SFB08']; 
              $rxno=$row2['RXNO'];   
              $sfb01=$row2['SFB01'];   
              $gem02=$row2['GEM02'];
              $msg= findcasewithsfb01($sfb01,$gem02,$erp_conn1,$erp_conn2);            

              if ( $rxno!=$oldrxno) {
                $i++;  
                $ii=$i;
                $rxno =$row2['RXNO'];
                $oldrxno=$row2['RXNO'];
                $indate=$row2['INDATE'];
                $duedate=$row2['DUEDATE'];               
              } else {     
                $rxno='';
                $ii='';
                $indate='';
                $duedate='';    
              }
                          
                                
              ?>
                  <tr bgcolor="#<?=$bgkleur;?>">
                      <td><?=$ii;?></td>      
                      <td><?=$indate;?></td>  
                      <td><?=$rxno;?></td>    
                      <td><?=$row2["OEA01"];?></td> 
                      <td><?=$duedate;?></td>
                      <td><?=$row2["IMA01"];?></td>  
                      <td><?=$row2["IMA02"];?></td>  
                      <td><?=$row2["SFB08"];?></td>  
                      <td><?=$faxout;?></td>   
                      <td><?=$faxin;?></td>      
                      <td><?=$msg;?></td>   
                    </tr>
              <?  
          } 
      }
      ?>
      <tr bgcolor="#<?=$bgkleur;?>">
        <td></td>      
        <td>Total</td>  
        <td colspan="5"></td>    
        <td><?=$total;?></td>  
        <td colspan="3"><?=$faxout;?></td>  
      </tr>   
          
</table>   