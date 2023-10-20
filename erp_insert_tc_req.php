<?
  session_start();
  $pagtitle = "業務部 &raquo; 傳真扣留作業v2"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_faxinv2.php");   
  date_default_timezone_set('Asia/Taipei');

  $soea = "select oea01 from oea_file where oea04 = 'E135001' and oea02=to_date('161217','yy/mm/dd') ";
  $erp_sqloea=oci_parse($erp_conn2,$soea);
  oci_execute($erp_sqloea); 
  while ($rowoea = oci_fetch_array($erp_sqloea, OCI_ASSOC)) {
      $oea01 = $rowoea['OEA01'];
      $sreq  = "insert into tc_req_file (tc_req001, tc_req002, tc_req003, tc_req004) values " .
               "('$oea01', '1', '空间不够,若对咬真牙修对咬，若支台齿是种植体可直接修支台齿； 若对咬假牙不能修支台齿，请送回传真。','1')";

      $erp_sqlreq1=oci_parse($erp_conn1,$sreq);
      $rs1=oci_execute($erp_sqlreq1);   

      $erp_sqlreq2=oci_parse($erp_conn2,$sreq);
      $rs2=oci_execute($erp_sqlreq2); 
  }