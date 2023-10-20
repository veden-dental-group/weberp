<?php
	session_start();
	$pagetitle = "報關組 &raquo; 每日出貨類別組數";
	include("_data.php");
	include("_erp.php");
	//auth("erp_weighttypecount.php");  

	if (is_null($_GET['bdate'])) {
		$bdate =  date('Y-m').'-01';
	} else {
		$bdate=$_GET['bdate'];
	}     

	if (is_null($_GET['edate'])) {
		$edate =  date('Y-m-d');
	} else {
		$edate=$_GET['edate'];
	}                                
	if ($_GET["submit"]=="匯出") {      
		error_reporting(E_NONE);  
		require_once 'classes/PHPExcel.php'; 
		require_once 'classes/PHPExcel/IOFactory.php';  
		$objReader = PHPExcel_IOFactory::createReader('Excel5');
		$objPHPExcel = $objReader->load("templates/erp_weighttypecount.xls");    

		$s2="select to_char(tc_oga002,'yyyy-mm-dd') tc_oga002, tc_ogb007, count(*) cases from tc_oga_file, tc_ogb_file " .
			"where tc_oga001=tc_ogb001 and tc_oga002 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') " .
		    "group by tc_oga002, tc_ogb007 order by tc_oga002, tc_ogb007 ";  
		//工單的日期以訂單的order date為到貨日期       
		$erp_sql2 = oci_parse($erp_conn1,$s2 );
		oci_execute($erp_sql2);  
		$y=1;           
		$oldate='';
		while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
			if ($olddate!=$row2['TC_OGA002']) {
				$y++;
				$olddate=$row2['TC_OGA002'];				
		  		$objPHPExcel->setActiveSheetIndex(0)
		       	   			->setCellValue('A'.($y), $olddate) ;
			}        
		    switch ($row2['TC_OGB007']) {
		    	case '9211':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('B'.($y), $row2['CASES']) ;
		       		break;
		    	case '9212':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('C'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9213':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('D'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9214':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('E'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9215':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('F'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9216':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('G'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9217':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('H'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9218':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('I'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9219':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('J'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9220':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('K'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9221':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('L'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9222':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('M'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9223':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('N'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9224':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('O'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9225':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('P'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9226':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('Q'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9227':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('R'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9228':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('S'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9229':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('T'.($y), $row2['CASES']) ;
		       		break;	
		    	case '9230':
		    		$objPHPExcel->setActiveSheetIndex(0)
		       		    ->setCellValue('U'.($y), $row2['CASES']) ;
		       		break;	
		    }
		}                                                                                                                        
		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(0);

		// Redirect output to a client’s web browser (Excel5)
		header('Content-Type: application/vnd.ms-excel');  
		header('Content-Disposition: attachment;filename="erp_weighttypecount.xls"');    
		header('Cache-Control: max-age=0'); 
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
		$objWriter->save('php://output'); 
		exit;    
	}  
    
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>期間內每日出貨類別組數合計 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee" width="60">出貨區間:</td>
          <td> 
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
            <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">  &nbsp; &nbsp; &nbsp; &nbsp; 
            <input type="submit" name="submit" id="submit" value="匯出">         
          </td>
      </tr>
    </table>
  </div>
</form> 
