<?php
  session_start();
  $pagetitle = "資材部 &raquo; 錄入鑄造後重量";
  include("_data.php");     
  //auth("erp_metal_add.php");

  if ($_POST["action"] == "save") { 
                                                         
      for ($x=1; $x<=$_POST["howmany"];$x++){            
            $queryp = "insert into erp_metal_add ( metal, bdate, ticketno, weight1, weight2, memo) values (       
                      '" . safetext($_POST["metal".$x])  . "',
                      '" . safetext($_POST["bdate".$x])      . "', 
                      '" . safetext('A311-'.$_POST["ticketno".$x])     . "', 
                      '" . safetext($_POST["weight1".$x])    . "', 0 , 
                      '" . safetext($_POST["memo".$x])       . "')"; 
            $resultp = mysql_query($queryp) or die ('47 erp_metal_weight_add error!!'.mysql_error()); 
      }    
      msg('資料儲存完畢!!');
	    forward('erp_nf_add.php' );   
  }
  
    
  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');
  
  Function showCases($v){      
      $o=new xajaxResponse();
      $msg = "";

      if ($v["bdate"]==""){
          $msg .= "請輸入日期. \n";        
      } 
      if ($v["metal"]==""){
          $msg .= "請輸入金屬 -- nf \n";        
      }
      if (substr($v["ticketno"],0,4)=="A311"){
          $msg .= "工單號不用輸入A311- \n";        
      }
      if (strlen($v["ticketno"])!="10"){
          $msg .= "工單號長度為10個字元 \n";        
      }
      if ($v["weight1"]==""){
          $msg .= "請輸入牙重 \n" ;        
      }  
      if ($v["weight1"]==0){
          $msg .= "請輸入牙重 \n" ;        
      }  
      if ($msg != ""){                           
          $o->alert($msg);         
      } else {  
          $x         = $v["howmany"];
          $bdate     = $v["bdate"];
          $metal     = $v["metal"];  
          $ticketno  = $v["ticketno"];  
          $weight1   = $v["weight1"];    
          $memo      = $v["memo"];                   
                
          if ($x<1) {                   
              $title  = "<input type='text' value='日期'        size='25' readonly='true'>";
              $title .= "<input type='text' value='金屬'        size='25' readonly='true'>";      
              $title .= "<input type='text' value='工單'        size='25' readonly='true'>";  
              $title .= "<input type='text' value='鑄造後重量'  size='25' readonly='true'>";   
              $title .= "<input type='text' value='備註'        size='25' readonly='true'>";                             
              $title .= "<br>";
              $o->append("insertarea", "innerHTML", $title); 
              $x=1;
          } else {
              $x ++;
          }  
             
          $case  = "<input type='text' name='bdate"     . $x . "' id='bdate"      . $x . "' size='25' value='" . $bdate . "'>"; 
          $case .= "<input type='text' name='metal"     . $x . "' id='metal"      . $x . "' size='25' value='" . $metal . "'>";    
          $case .= "<input type='text' name='ticketno"  . $x . "' id='ticknetno"  . $x . "' size='25' value='" . $ticketno . "'>"; 
          $case .= "<input type='text' name='weight1"   . $x . "' id='weight1"    . $x . "' size='25' style='text-align:right' value='" . $weight1 . "' onkeypress='return numberOnly(event,\"weight" . $x . "\")' >";   
          $case .= "<input type='text' name='memo"       . $x . "' id='memo" . $x . "' size='25' value='" . $memo . "'>";             
          $case .= "<br>";     
          $o->append("insertarea", "innerHTML", $case );                                                          
            
          $o->assign("ticketno","value",  "" );   
          $o->assign("weight1","value",  "0" );     
          $o->assign("howmany", "value", $x);
      }
      return $o; 
  } 
  
  Function checkrxno($v){      
      $o=new xajaxResponse();   
      $ticketno  = $v['ticketno'];  
      if ($ticketno!=""){
          $erp_db_host3 = "topprod";
          $erp_db_user3 = "vd110";
          $erp_db_pass3 = "vd110";  
          $erp_conn3 = oci_connect($erp_db_user3, $erp_db_pass3, $erp_db_host3,'AL32UTF8'); 
          $ssfb="select sfbud02 from sfb_file where sfb01='$ticketno'";
          $erp_sqlsfb = oci_parse($erp_conn3,$ssfb );
          oci_execute($erp_sqlsfb);  
          $rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC);
          $rxno=$rowsfb['SFBUD02'];
          if (is_null(!$result)) {
              $o->alert("164 查無此工單的 RX NO!!" . mysql_error());   
          } else {
              $o->assign("rxno","value", $rxno);                                                      
          }                 
      }   
      return $o;
  }        
      
  $xajax->register(XAJAX_FUNCTION,'checkrxno');     
  $xajax->register(XAJAX_FUNCTION,'showCases');  
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True;             
  
  $showodata = "請輸入以下資料以新增 鑄造後重量." ;            
  $bdate = is_null($_GET["bdate"]) ? date('Y-m-d'):$_GET["bdate"];    
  $metal = is_null($_GET["metal"]) ? 'nf':$_GET["metal"];                    
  include("_header.php");
?>                                  
  
<form name="form1" method="post" action="<?=$PHP_SELF;?>" > 
 
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr><td bgcolor="#666666">
        <table width="100%"  border="0" cellpadding="3" cellspacing="1" style="border: 1px">
          <tr>            
            <td class="inputw">日期:
                <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()">   
            </td>   
            <td class="inputw">金屬:       <input type="text" name="metal"    id="metal"    size="5"  value=<?=$metal;?>>  
            <td class="inputw">工單:A311-  <input type="text" name="ticketno" id="ticketno" size="15" ></td>    
            <td class="inputw">鑄造後重量: <input type="text" name="weight1"  id="weight1"  size="5"  value='0' style="text-align:right" onkeypress="return numberOnly(event,'weight1')"></td>        
            <td class="inputw">備註:      <input type="text" name="memo"     id="memo"     size="15">  
            <td class="inputw">           <input type="button" value="加" onclick="xajax_showCases(xajax.getFormValues('form1'))"></td>  
          </tr>
        </table>
      </td></tr>
  </table>          
       
  <table width="100%" border="0" cellspacing="1" cellpadding="1" style="BORDER: 1px solid #666666">     
    <tr><td> 
      <input type="hidden" name="howmany" id="howmany">   
      <div id="insertarea">  </div> 
    </td></tr>                                            
  </table> 
        
  <table>
    <tr>
      <td><input type="hidden" name="action" value="save">  
          <input type="submit" name="Submit" value="存檔">       
      </td>
    </tr>  
  </table>
</form>
