<?
  session_start();
  $pagtitle = "報關組 &raquo; 秤重單未分配到報關金額"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_weightnomoney.php");
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
                ->setCellValue('A1', $thisyear.' 秤重單未分配到報關金額');       
                
    $y=1;                    
    $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.$y ,  '秤重日期');
    $s1= "select distinct to_char(tc_oga002,'yyyy-mm-dd') tc_oga002 from tc_oga_file, tc_ogb_file where tc_oga001=tc_ogb001 and tc_ogb008 is null order by tc_oga002";
    $erp_sql1 = oci_parse($erp_conn1,$s1 );
    oci_execute($erp_sql1);  
    $y=2; 
    while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) { 
          $objPHPExcel->setActiveSheetIndex(0)       
                      ->setCellValue('A'.$y ,  $row1['TC_OGA002']) ;
          $y++;     
    }                      
                
    $objPHPExcel->getActiveSheet()->setTitle('秤重單未分配到報關金額');
                                                            
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: application/vnd.ms-excel');  
    header('Content-Disposition: attachment;filename="'. $bdate. '_'.$edate.'_weightnomoney.xls"');   
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
<p>有秤重分類卻無報關分類 </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            秤重期間:   
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
    <th>秤重日期</th>       
  </tr>
  <?          
      $s1= "select distinct to_char(tc_oga002,'yyyy-mm-dd') tc_oga002 from tc_oga_file, tc_ogb_file where tc_oga001=tc_ogb001 and tc_ogb008 is null order by tc_oga002"; 
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);   
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
      ?>    
	      <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
          <td><?=$row1["TC_OGA002"];?></td>                           
    <?  
      }   
  ?>     
  <tr><td colspan="2">若有資料, 請至 csft997 中去重新攤當天的報關金額</td></tr>       
</table>  
</form>
