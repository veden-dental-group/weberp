<?php
    session_start();
    //$pagetitle = "資訊部 &raquo; 全部客戶的全部價格";
    include("_data.php");
    include("_erpdata.php");
    //auth("erp_allclientsprices.php");   
  
   // if ($_GET["submit"]=="Excel") {        
		$filename="templates/allclientsprices.xls";   
		error_reporting(E_ALL);  

		require_once 'classes/PHPExcel.php'; 
		require_once 'classes/PHPExcel/IOFactory.php';  
		$objReader = PHPExcel_IOFactory::createReader('Excel5'); 
		$objPHPExcel = $objReader->load("$filename");  
		         
        $s1="select occ01, occ02 from occ_file order by occ01 ";                                           
		$erp_sql1 = oci_parse($erp_conn1,$s1 );
        oci_execute($erp_sql1);
        $sheet=1; 
        $x=1;        
        while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
            //echo $row1['OCC01'];        
        	if (($x % 100)== 1) {  
        	    $s = $objPHPExcel->setActiveSheetIndex($sheet);   
				//A3,B3 填品代/品名
				$s2="select ima01, ima02 from ima_file where ima08='M' order by ima01";
				$erp_sql2 = oci_parse($erp_conn2,$s2 );
		        oci_execute($erp_sql2);
		        $y=3;
		        while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) { 
		        	$s->setCellValue('A'.$y, $row2['IMA01'])
              		  ->setCellValue('B'.$y, $row2['IMA02']);
              		//$s->getStyle('A'.$y)->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		            //$s->getStyle('A'.$y)->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		            //$s->getStyle('A'.$y)->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		            //$s->getStyle('A'.$y)->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
		            //$s->getStyle('B'.$y)->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		            //$s->getStyle('B'.$y)->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		            //$s->getStyle('B'.$y)->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		            //$s->getStyle('B'.$y)->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);  
              		$y++; 
		        }	
				$sheet++;  
				$x=1;                
            } 

            $z= $x + 66 ;
            if ( $z > 90){ //超過90以後的 都要加一個文字
                $a = chr(floor(($z-65)/26)+64);
                $b = chr((($z-65) % 26)+65) ;
            } else {
                $a="";
                $b=chr($z); 
            }  

            $s->setCellValue($a.$b.'1', $row1['OCC01'])
              ->setCellValue($a.$b.'2', $row1['OCC02']);

            $s3="select ima01, xmf07 from " .
				" (select occ01, occ02, occ44, ima01, ima02 from ima_file, occ_file where ima08='M' and  occ01='". $row1['OCC01']. "' ) " .
				" left join xmf_file " .
				" on occ44=xmf01 and ima01=xmf03 " .
				" order by ima01";
			$erp_sql3 = oci_parse($erp_conn1,$s3 );
	        oci_execute($erp_sql3);
	        $y=3;
	        while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) { 
	            $s->setCellValue($a.$b.$y, $row3['XMF07']); 
	            $s->getStyle($a.$b.$y) ->getNumberFormat() ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00 );  
	            //$s->getStyle($a.$b.$y)->getBorders() ->getTop() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            //$s->getStyle($a.$b.$y)->getBorders() ->getBottom() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            //$s->getStyle($a.$b.$y)->getBorders() ->getLeft() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            //$s->getStyle($a.$b.$y)->getBorders() ->getRight() ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	            $y++;
	        }
            $x++;
        }
      
        $objPHPExcel->setActiveSheetIndex(0);             
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5'); 
        $filename='report/allclientsprices.xls';
        $objWriter->save($filename);   
        exit;
 /* }  
    
    //for xajax
  //require ('xajax/xajax_core/xajax.inc.php');
  //$xajax = new xajax();
  //$xajax->configure('javascript URI', 'xajax/');
    
  //$xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>所有客戶所有產品價格列印 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">   
        <input type="submit" name="submit" id="submit" value="Excel">  &nbsp;&nbsp;   &nbsp;&nbsp;      
      </td></tr>
    </table>
  </div>
</form>
*/
