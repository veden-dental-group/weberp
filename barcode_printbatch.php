<?
  session_start();
  $pagtitle = "IT作業 &raquo; 批次列印工號"; 
  include("_data.php");
  auth("barcode_printbatch.php");
  
  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');

  Function prePrintBarcodes($v, $x){              
      $o=new xajaxResponse(); 
      $js= "printBarcode('" . $v . "','". $x . "')";  
      //$o->alert($js);
      $o->script($js);         
      return $o; 
  }  
          
  $xajax->register(XAJAX_FUNCTION,'prePrintBarcodes'); 
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True; 
 
  if (is_null($_GET['x'])) {
    $x=110;
  } else {
    $x=$_GET['x'];
  }
  include("_header.php");
?>
<script language='JavaScript'>

function printBarcode(v,x) { 
  
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
     htzcmd.PrnQty(1);
     htzcmd.PrnEndPage();
     htzcmd.PrnClose();   
    }
  }
  catch(e)
  { alert(e+ " "+ v); } 
}

</script>
<OBJECT id=htzcmd classid=clsid:5E566789-F012-4B54-8882-38E3BD3D847B></OBJECT>  
<link href="css.css" rel="stylesheet" type="text/css">
<p>批次列印員工工號條碼. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">部門:
            <select name="gem01" id="gem01">  
              <option value="">請選擇一個部門</option>  
              <?
                $s= "select gem01, gem02 from gem_file order by gem01";
                $erp_sql = oci_parse($erp_conn, $s);
                oci_execute($erp_sql);                        
                while ($row1 = oci_fetch_array($erp_sql, OCI_ASSOC)) {
                   echo "<option value=" . $row1["GEM01"];
                  if ($_GET["gem01"] == $row1["GEM01"]) echo " selected";
                  echo ">" . $row1['GEM01'] . "." . $row1["GEM02"] . "</option>";
                }       
              ?>
            </select>
            &nbsp;  &nbsp;     
            <input type="submit" name="Submit2" value="查詢" />&nbsp;  &nbsp;      
            X軸位置: 
            <input name="x" id="x" value="<?=$x;?>" size="5">
            </div>
          </td>          
        </tr>
    </table>
  </div>
</form>

<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>部門代號</th>
    <th>部門名稱</th>
    <th>員工代號</th>
    <th>員工姓名</th>  
    <th width="16">&nbsp;</th>
  </tr>
  <?
  if (is_null($_GET['gem01'])) die();
  $gem01=substr($_GET['gem01'],0, strpos($_GET['gem01'],'0'));
  $gem01=$_GET['gem01'];
  $s= "select gem01, gem02, gen01, gen02 from gem_file, gen_file where gem01 ='$gem01' and gem01=gen03";  
  $erp_sql = oci_parse($erp_conn, $s);
  oci_execute($erp_sql);                        
  while ($row1 = oci_fetch_array($erp_sql, OCI_ASSOC)) {   
    $bgcolor = "ffffff";      
    ?>    
	    <tr bgcolor="#<?=$bgcolor;?>"> 
		      <td><img src="i/arrow.gif" width="16" height="16"> </td> 
          <td><?=$row1["GEM01"];?></td>
          <td><?=$row1["GEM02"];?></td>
			    <td><?=$row1["GEN01"];?></td>
			    <td><?=$row1["GEN02"];?></td>
          <td width=16><input type="Button" name="Submit" value="P" onclick="xajax_prePrintBarcodes('<?=$row1['GEN01'];?>','<?=$_GET['x'];?>')"> </td>  
     </tr> 
  <?  
	}
	?>   
</table>  
</form>
