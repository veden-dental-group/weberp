<?
  session_start();
  $pagtitle = "廠務部 &raquo; 每日到貨CASE追蹤"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_caseindailytrace.php");
  if (is_null($_GET['tdate'])) {
    $tdate = date('Y-m-d');
  } else {
    $tdate=$_GET['tdate'];
  } 
  
  if (is_null($_GET['maker'])) {
    $maker = '';
  } else {
    $maker=$_GET['maker'];
  }  
  
  if ($_GET['maker']=='') {
      $makerfilter='';
  } else {    
      $makerfilter=" and sfb82='". $_GET['maker'] . "'";
  }
  
  if ($_GET['issent']=='Y') {                  
  } else {
      $issentfilter=" where ogb03 is null ";       
  }
  
  $issentfilter=" ";
  
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
                ->setCellValue('A1', $thisyear.' 每日到貨CASE追蹤');       
                
    $y=1;                    
    $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.$y ,  '日期')           
                      ->setCellValue('B'.$y ,  '製處')             
                      ->setCellValue('C'.$y ,  '工序代號') 
                      ->setCellValue('D'.$y ,  '工序名稱') 
                      ->setCellValue('E'.$y ,  '到貨數(組)') 
                      ->setCellValue('F'.$y ,  '試戴進(顆)')  
                      ->setCellValue('G'.$y ,  'POST(顆)')     
                      ->setCellValue('H'.$y ,  '磁牙到貨(顆)')  
                      ->setCellValue('I'.$y ,  '鋼牙到貨(顆)')  
                      ->setCellValue('J'.$y ,  '刷進數(組)') 
                      ->setCellValue('K'.$y ,  '到貨未刷(組)') 
                      ->setCellValue('L'.$y ,  '刷出數(組)') 
                      ->setCellValue('M'.$y ,  '刷進未刷出(組)') 
                      ->setCellValue('N'.$y ,  '出貨數(組)') 
                      ->setCellValue('O'.$y ,  '到貨未出(組)');    
    $s1="select to_char(oea02,'yy/mm/dd') oea02, sfb82, gem02, ecb06, ecb17, sum(casetotal) casetotal, sum(type1) type1, sum(type2) type2, sum(type3) type3, sum(type4) type4, sum(type9) type9, sum(casein) casein, sum(caseout) caseout, sum(casesent) casesent from " .
              "(select oea02, sfb82, gem02, ecb06, ecb17, decode(imaud02,'1', sfb08, 0) type1, decode(imaud02,'2', sfb08, 0) type2, decode(imaud02,'3', sfb08, 0) type3, decode(imaud02,'4', sfb08, 0) type4, decode(imaud02,'9', sfb08, 0) type9, 1 casetotal, decode(tc_srg007, null, 0, 1) casein , decode(tc_srg010, null, 0, 1) caseout, decode(ogb03, null, 0, 1) casesent from " .
                  "(select  oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, imaud02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011, ogb03 from " .
                      "(select  oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, imaud02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                          "(select  oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, imaud02, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                              "(select oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, substr(imaud02,1,1) imaud02 " .
                              "from sfb_file, gem_file, ima_file, oea_file " .
                              "where sfb22=oea01 $makerfilter and oea02=to_date('$tdate','yy/mm/dd') " .
                              "and sfb82=gem01 and sfb05=ima01) " .   
                      "left join tc_srg_file  on  sfb01=tc_srg001 ) , ecb_file where sfb05=ecb01 and tc_srg004=ecb03 ) " .
                  "left join ogb_file on sfb22=ogb31 and sfb221=ogb32 $issentfilter) ".
              ") " . 
          "group by oea02, sfb82, gem02, ecb06, ecb17 " . 
          "order by oea02, sfb82, gem02, ecb06, ecb17 " ;       
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);  
    $y=2;                                  
    while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) { 
          $type23=$row1['TYPE2']+ $row1['TYPE3'] ;         
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.$y ,  $row1['OEA02']) 
                      ->setCellValue('B'.$y ,  $row1['SFB82'].' ' .$row1['GEM02']) 
                      ->setCellValue('C'.$y ,  $row1['ECB06']) 
                      ->setCellValue('D'.$y ,  $row1['ECB17']) 
                      ->setCellValue('E'.$y ,  $row1['CASETOTAL'])  
                      ->setCellValue('F'.$y ,  $row1['TYPE1'])
                      ->setCellValue('G'.$y ,  $row1['TYPE9'])
                      ->setCellValue('H'.$y ,  $type23)
                      ->setCellValue('I'.$y ,  $row1['TYPE4'])
                      ->setCellValue('J'.$y ,  $row1['CASEIN']) 
                      ->setCellValue('K'.$y ,  $row1['CASETOTAL']-$row1['CASEIN']) 
                      ->setCellValue('L'.$y ,  $row1['CASEOUT']) 
                      ->setCellValue('M'.$y ,  $row1['CASEIN']-$row1['CASEOUT']) 
                      ->setCellValue('N'.$y ,  $row1['CASESENT']) 
                      ->setCellValue('O'.$y ,  $row1['CASETOTAL']-$row1['CASESENT']);   
          $y++;                                                                            
    }                      
                
    $objPHPExcel->getActiveSheet()->setTitle('每日到貨CASE追蹤');
                                                            
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
<link href="css.css" rel="stylesheet" type="text/css">
<p>每日到貨CASE追蹤 </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            製處:
            <select name="maker" id="maker">  
              <option value="">全部</option>  
              <?
                $q2 = "select mid, mname from maker order by mid";
                $r2 = mysql_query($q2) or die ('51 maker error!!');
                while ($rr2 = mysql_fetch_array($r2)) {
                   echo "<option value=" . $rr2["mid"];
                  if ($_GET["maker"] == $rr2["mid"]) echo " selected";
                  echo ">" . $rr2["mid"] . ' '.$rr2['mname'] . "</option>";
                }       
              ?>
            </select>  &nbsp;&nbsp;&nbsp;&nbsp; 
            到貨日期:   
            <input name="tdate" type="text" id="tdate" size="12" maxlength="12" onfocus="new WdatePicker()" value=<?=$tdate;?>> &nbsp; &nbsp;        
      <!--      <input type="checkbox" name="issent" id="issent" value="Y" <? if ($_GET['issent']=='Y') echo " checked";?>>含已出貨  &nbsp; &nbsp;  -->                                     
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
    <th>日期</th> 
    <th>製處</th> 
    <th>工序代號</th>     
    <th>工序名稱</th>   
    <th>到貨數(組)</th>    
    <th>試戴進(顆)</th>   
    <th>POST(顆)</th>  
    <th>磁牙到貨(顆)</th>   
    <th>鋼牙到貨(顆)</th>   
    <th>刷進數(組)</th> 
    <th>到貨未刷(組)</th>   
    <th>刷出數(組)</th>   
    <th>刷進未刷出(組)</th>  
    <th>出貨數(組)</th>   
    <th>到貨未出(組)</th>                                                                                         
  </tr>
  <?          
      $s1="select to_char(oea02,'yy/mm/dd') oea02, sfb82, gem02, ecb06, ecb17, sum(casetotal) casetotal, sum(type1) type1, sum(type2) type2, sum(type3) type3, sum(type4) type4, sum(type9) type9, sum(casein) casein, sum(caseout) caseout, sum(casesent) casesent from " .
              "(select oea02, sfb82, gem02, ecb06, ecb17, decode(imaud02,'1', sfb08, 0) type1, decode(imaud02,'2', sfb08, 0) type2, decode(imaud02,'3', sfb08, 0) type3, decode(imaud02,'4', sfb08, 0) type4, decode(imaud02,'9', sfb08, 0) type9, 1 casetotal, decode(tc_srg007, null, 0, 1) casein , decode(tc_srg010, null, 0, 1) caseout, decode(ogb03, null, 0, 1) casesent from " .
                  "(select  oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, imaud02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011, ogb03 from " .
                      "(select  oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, imaud02, ecb06, ecb17, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                          "(select  oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, imaud02, tc_srg004, tc_srg005, tc_srg007, tc_srg008, tc_srg010, tc_srg011 from " .
                              "(select oea02, sfb01, sfb22, sfb221, sfb08, sfb82, gem02, sfb05, substr(imaud02,1,1) imaud02 " .
                              "from sfb_file, gem_file, ima_file, oea_file " .
                              "where sfb22=oea01 $makerfilter and oea02=to_date('$tdate','yy/mm/dd') " .
                              "and sfb82=gem01 and sfb05=ima01) " .   
                      "left join tc_srg_file  on  sfb01=tc_srg001 ) , ecb_file where sfb05=ecb01 and tc_srg004=ecb03 ) " .
                  "left join ogb_file on sfb22=ogb31 and sfb221=ogb32 $issentfilter) ".
              ") " . 
          "group by oea02, sfb82, gem02, ecb06, ecb17 " . 
          "order by oea02, sfb82, gem02, ecb06, ecb17 " ;       
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);   
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) { 
        $type23 = $row1["TYPE2"] + $row1["TYPE3"];
      ?>    
	      <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
          <td><?=$row1["OEA02"];?></td>  
          <td><?=$row1["SFB82"].' '.$row1["GEM02"];?></td>
          <td><?=$row1["ECB06"];?></td> 
          <td><?=$row1["ECB17"];?></td>  
          <td><?=$row1["CASETOTAL"];?></td>   
          <td><?=$row1["TYPE1"];?></td>   
          <td><?=$row1["TYPE9"];?></td>   
          <td><?=$type23;?></td>   
          <td><?=$row1["TYPE4"];?></td> 
          <td><?=$row1["CASEIN"];?></td> 
          <td style="color:#aa0000"><?=$row1['CASETOTAL']-$row1["CASEIN"];?><a href="erp_caseindailytrace1.php?maker=<?=$maker;?>&tdate=<?=$tdate;?>&ecb06=<?=$row1['ECB06'];?>&issent=<?=$_GET['issent'];?>&isecb06=<?=$_GET['isecb06'];?>" target="_blank"><img src="i/info.gif" width="16" height="16" border="0" alt="查詢"></td>   
          <td><?=$row1["CASEOUT"];?></td>  
          <td style="color:#dd0000"><?=$row1["CASEIN"]-$row1['CASEOUT'];?><a href="erp_caseindailytrace2.php?maker=<?=$maker;?>&tdate=<?=$tdate;?>&ecb06=<?=$row1['ECB06'];?>&issent=<?=$_GET['issent'];?>&isecb06=<?=$_GET['isecb06'];?>" target="_blank"><img src="i/info.gif" width="16" height="16" border="0" alt="查詢"></a></td>  
          <td><?=$row1["CASESENT"];?></td>    
          <td style="color:#ff0000"><?=$row1["CASETOTAL"]-$row1['CASESENT'];?><a href="erp_caseindailytrace3.php?maker=<?=$maker;?>&tdate=<?=$tdate;?>&ecb06=<?=$row1['ECB06'];?>&issent=<?=$_GET['issent'];?>&isecb06=<?=$_GET['isecb06'];?>" target="_blank"><img src="i/info.gif" width="16" height="16" border="0" alt="查詢"></a></td>       
      <?  
      }   
  ?>            
</table>  
</form>
