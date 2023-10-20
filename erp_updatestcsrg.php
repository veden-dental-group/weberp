<?
  session_start();
  include("_data.php");
  include("_erp.php");
  //更新開單日期 報工日期
        
  $ssfe="select tc_srg001, tc_srg004,  tc_srg007, tc_srg008,tc_srg010,tc_srg011, sfb38, sfb81 " .
        "from tc_srg_file,sfb_file where tc_srg001=sfb01";
  $erp_sqlsfe = oci_parse($erp_conn1,$ssfe );
  oci_execute($erp_sqlsfe); 
  while ($rowsfe = oci_fetch_array($erp_sqlsfe, OCI_ASSOC)) {    
      $srg001=$rowsfe['TC_SRG001'];        # 工單號
      $srg004=$rowsfe['TC_SRG004'];        # 項次
      $srg007=$rowsfe['TC_SRG007'];        # 進站日期
      $srg008=$rowsfe['TC_SRG008'];        # 進站時間
      $srg010=$rowsfe['TC_SRG010'];        # 出站日期
      $srg011=$rowsfe['TC_SRG011'];        # 出站時間
      $sfb38 =$rowsfe['SFB38'];            # 結案日  不動
      $sfb81 =$rowsfe['SFB81'];            # 開單日

      #出站>結案日 出站=結案日
      if ($srg010 > $sfb38) $srg010=$sfb38;      
            
      #進站>出站 進站=出站
      if ( $srg007 > $srg010 ) {
          $srg007=$srg010;        
      } 
      
      #進日=出日 進時>出時  進站-1      
      if ($tc_srg007=$tc_srg010 and $tcsrg008 > $tcsrg011) {
          $tcsrg007 -- ;
      }
      
      #開單日=進站
      
           
      
  } 
?>
