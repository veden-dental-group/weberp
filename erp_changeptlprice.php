<?
  session_start();
  $pagtitle = "帳單組 &raquo; 更改PTL價格"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_changeptlprice.php");
   
  if (is_null($_GET['sdate'])) {
    $sdate = Date('Y-m-d');
  } else {
    $sdate=$_GET['sdate'];
  }       
  
    
  if ($_GET["Submit"] == "Update") {
      //先更新單價
      $sofb="update tc_ofb_file set tc_ofb13 = (select xmf07 from xmf_file where xmf01='PTL' and xmf03=tc_ofb04) " .
            "where tc_ofb01 in ( select tc_ofa01 from tc_ofa_file where tc_ofa03='E129000' " .
            "and tc_ofa02= to_date('$sdate','yy/mm/dd'))";  
      $erp_sqlofb = oci_parse($erp_conn2,$sofb);
      oci_execute($erp_sqlofb); 

      //更新總價
      $sofb="update tc_ofb_file set tc_ofb14=tc_ofb12*tc_ofb13, tc_ofb14t=tc_ofb12*tc_ofb13 " .
            "where tc_ofb01 in ( select tc_ofa01 from tc_ofa_file where tc_ofa03='E129000' " .
            "and tc_ofa02=to_date('$sdate','yy/mm/dd'))";  
      $erp_sqlofb = oci_parse($erp_conn2,$sofb);
      oci_execute($erp_sqlofb); 
      
      msg('價格更新完畢!!');          
      forward("erp_changeptlprice.php");                                                              
  }
  
  include("_header.php");
?>

<link href="css.css" rel="stylesheet" type="text/css">
<p>更新 PTL Invoice價格 </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">    
            Date:
            <input name="sdate" type="text" id="sdate" size="12" maxlength="12" value=<?=$sdate;?> onfocus="new WdatePicker()">&nbsp; &nbsp;            
            <input type="submit" name="Submit" value="Update" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>
                 