<?
  session_start();
  include("_data.php");
  include("_erp.php");
  
  //因VMI有打勾無法立帳的資料刷成 N
  $svmi="update rvb_file set rvb89='N' where rvb89 is null or rvb89='Y'";
  $erp_sqlvmi = oci_parse($erp_conn1,$svmi );
  oci_execute($erp_sqlvmi); 
  
  $svmi="update rvv_file set rvv89='N' where rvv89 is null or rvv89='Y'";
  $erp_sqlvmi = oci_parse($erp_conn1,$svmi );
  oci_execute($erp_sqlvmi);
  
  $svmi="update pmn_file set pmn89='N' where pmn89 is null or pmn89='Y'";
  $erp_sqlvmi = oci_parse($erp_conn1,$svmi );
  oci_execute($erp_sqlvmi);

  msg('VMI更新成功!!');
  forward("main.php");   
?>
