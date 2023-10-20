<?
  session_start();
  include("_data.php");
  include("_erp.php");
  
  //取出沒有備料記錄的工單           
  $ssfe="select sfe01, sfe07, sfe14, sfe17, sfe27, sum(sfe16) sfe16 from 
          (
          select * from sfe_file left outer join sfa_file on sfe01=sfa01 and sfe07=sfa03 and sfe14=sfa08 and sfe17=sfa12 and sfe27=sfa27 
          ) 
          where sfa01 is null group by sfe01, sfe07, sfe14, sfe17, sfe27 order by sfe01";
  $erp_sqlsfe = oci_parse($erp_conn1,$ssfe );
  oci_execute($erp_sqlsfe); 
  while ($rowsfe = oci_fetch_array($erp_sqlsfe, OCI_ASSOC)) {    
      $sfe01=$rowsfe['SFE01'];               //工單號
      $sfe14=$rowsfe['SFE14'];               //製程      
      
      $ssfb="select sfb05, sfb08 from sfb_file where sfb01='$sfe01' "; 
      $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
      oci_execute($erp_sqlsfb); 
      $rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC);
      $sfb05=$rowsfb['SFB05'];                     // 品代
      $sfb08=floatval($rowsfb['SFB08']);                     //顆數
      
      $sbmb="select bmb06 from bmb_file where bmb01='$sfb05' and bmb09='$sfe14' "; 
      $erp_sqlbmb = oci_parse($erp_conn1,$sbmb);
      oci_execute($erp_sqlbmb); 
      $rowbmb = oci_fetch_array($erp_sqlbmb, OCI_ASSOC);     //BOM用量
      $bmb06=floatval($rowbmb['BMB06']);
      
      $sfa02=1;
      $sfa03=$rowsfe['SFE07'];
      $sfa04=$sfb08 * $bmb06;
      $sfa06=$rowsfe['SFE16'];
      $sfa08=$rowsfe['SFE14'];
      $sfa12='G';
      $sfa16=$rowbmb['BMB06'];
      $sfa27=$rowsfe['SFE27'];
      $sfa29=$rowsfb['SFB05'];
      
      $ssfa="insert into sfa_file values('$sfe01',$sfa02,'$sfa03',$sfa04,$sfa04,$sfa06,0,0,0,0,0,0,0,'$sfa08',0,' ','N','$sfa12', 1, '$sfa12', 1, $sfa16, $sfa16, 0, 0, '$sfa27',1, '$sfa29', ' ',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ', 0, 'Y','N',' ','Frank','','','','',0,0,0,0,0,0,'','','','','VD110','VD100',' ',0)";
      $erp_sqlsfa = oci_parse($erp_conn1,$ssfa);
      oci_execute($erp_sqlsfa);        
      
  } 
?>
