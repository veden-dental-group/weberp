<?
  session_start();
  $pagtitle = "資材部 &raquo; 更改已採購量"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_changepml.php");
  

  if ($_POST["action"] == "save") { 
      $newpml21 =  $_POST['newpml21'] ;  
      if ($newpml21==0) {
          $soea= "update pml_file set pml21=0, pml16='1' where pml01='" . $_POST['pml01'] . "' and pml02='" . $_POST['pml02'] . "' ";  
      } else {
          $soea= "update pml_file set pml16='2', pml21= $newpml21 where pml01='" . $_POST['pml01'] . "' and pml02='" . $_POST['pml02'] . "' ";    
      }
      
      if ($_POST['conn']=='1') {
        $erp_sqloea = oci_parse($erp_conn1,$soea );   
      } else {
        $erp_sqloea = oci_parse($erp_conn2,$soea );   
      }                                         
      $rs=oci_execute($erp_sqloea);        
      msg("更新成功!!");
      forward("erp_changepml.php");
      exit;                                                              
  }
  
  include("_header.php");
?>

<link href="css.css" rel="stylesheet" type="text/css">
<p>更改採購量 </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            公司別:
            <select name="conn" id="conn"> 
              <option value='1' <? if($_GET['conn']=='1') echo " selected";?> > 義齒</option> 
              <option value='2' <? if($_GET['conn']=='2') echo " selected";?> > V-BEST</option>  
            </select> 
            請購單號:   
            <input name="pml01" type="text" id="pml01" size="20" maxlength="20"> &nbsp;  &nbsp;
            請購單項次:   
            <input name="pml02" type="text" id="pml02" size="5" maxlength="20"> &nbsp;  &nbsp;
            新已轉採購量:   
            <input name="pml21" type="text" id="pml21" size="5" value='0'> &nbsp;  &nbsp; 
            <input type="submit" name="Submit2" value="送出" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['Submit2'])) die ; ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>請購單號</th>
    <th>請購單項次</th> 
    <th>料號</th>   
    <th>品名規格</th>   
    <th>訂購量</th>   
    <th>原已轉採購量</th>
    <th>新已轉採購量</th>                                                                                 
  </tr>
  <?  
      $pmlfilter= " AND pml01='".$_GET['pml01'] ."' AND pml02='" . $_GET['pml02'] ."' ";
      $soea="select pml01, pml02, pml04, pml041, pml20, pml21 from pml_file " . 
            "where 1=1  " . $pmlfilter ; 
      if ($_GET['conn']=='1') {
        $erp_sqloea = oci_parse($erp_conn1,$soea );   
      } else {
        $erp_sqloea = oci_parse($erp_conn2,$soea );   
      }                                         
      oci_execute($erp_sqloea); 
      while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="conn" value="<?=$_GET['conn'];?>">  
              <input type="hidden" name="pml01" value="<?=$_GET['pml01'];?>"> 
              <input type="hidden" name="pml02" value="<?=$_GET['pml02'];?>"> 
              <input type="hidden" name="oldpml21" value="<?=$rowoea['PML21'];?>">  
              <input type="hidden" name="newpml21" value="<?=$_GET['pml21'];?>"> 
          </td>
			    <td><?=$rowoea["PML01"];?></td>
          <td><?=$rowoea["PML02"];?></td>   
          <td><?=$rowoea["PML04"];?></td> 
          <td><?=$rowoea["PML041"];?></td> 
          <td><?=$rowoea["PML20"];?></td>  
          <td><?=$rowoea["PML21"];?></td>  
          <td><?=$_GET["pml21"];?></td>       
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
