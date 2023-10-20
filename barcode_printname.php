<?php
session_start();

include("_data.php");
auth("barcode_printone.php");

//for xajax
require ('xajax/xajax_core/xajax.inc.php');
$xajax = new xajax();
$xajax->configure('javascript URI', 'xajax/');

Function prePrintBarcodes($v){      
    $o=new xajaxResponse();     
    $x=$v["xvalue"];
    $c=$v["context1"];         
    if ( $x % 8 ==0 ) {
      $printedbarcodes .= ("<br>" . '&nbsp;&nbsp;' . $orderno );
    } else {  
      $printedbarcodes .= ('&nbsp;&nbsp;' . $orderno) ;
    }
    $o->assign("printedBarcodes","innerHTML", $printedbarcodes); 
    $js= "printBarcode('" . $c . "','1','". $x . "')";  
    $o->script($js);
     
    //$o->alert("Barcode(s) printed!!"); 

   return $o; 
}  
        
$xajax->register(XAJAX_FUNCTION,'prePrintBarcodes'); 
$xajax->processRequest();

echo '<?xml version="1.0" encoding="UTF-8"?>'; 
$IsAjax = True;  

$pagetitle = "IT作業 &raquo; 列印單張條碼 ";    
include("_header.php");
?>

<script language=javascript type=text/javascript>
<!--

function printBarcode(v,q,x) { 
  try
   {  
      //                    PrintName   Darkness     Speed    Width   Height
     if ( htzcmd.PrnInit("SATO CX400",   "4A",         "3",     "50",    "15") == 0 )
     {                                                    
     htzcmd.PrnStartPage();    
     //                   X     Y    Barcode Type  Ratio   Space   Height  Rotate   Data
     htzcmd.PrnBarcode(x, "0010", "G",         "1:3",  "02",  "050",   "0",  v); 
    //            X        Y      FontName  Size  Bold   hight  width rotate  Data       
     htzcmd.PrnText(x,"0070","01",  "12",  "0",   "0",   "9",   "0", v);  
     //             Qty and Print
     htzcmd.PrnQty(q);
     htzcmd.PrnEndPage();
     htzcmd.PrnClose();   
    }
  }
  catch(e)
  { alert(e+ " "+ v); } 
}

</script>

<OBJECT id=htzcmd classid=clsid:5E566789-F012-4B54-8882-38E3BD3D847B></OBJECT> 
<link href="oos.css" rel="stylesheet" type="text/css">
<form action="<?=$_SERVER['PHP_SELF'];?>" method="post" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>   
        <td bgcolor="#eeeeee">
          <span align="left">內容:
            <input name="xvalue" type="text" id="xvalue" size="5" value='50'> 
            <input name="context1" type="text" id="context1" size="20" >  ( 派工單: 50  工號:100 )
          </span>
        </td>  
        <td bgcolor="#eeeeee">                             
          <input type="Button" name="Submit" value="Print" onclick="xajax_prePrintBarcodes(xajax.getFormValues('form1'))">   
        </td>   
      </tr>
    </table>
  </div>
  <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
    <tr>
      <td bgcolor="#eeeeee">
        <div align="left" id="printedbarcodes">
        </div>
      </td>
    </tr>   
  </table>
</form>
