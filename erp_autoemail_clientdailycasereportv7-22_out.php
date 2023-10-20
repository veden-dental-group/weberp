<?php
  date_default_timezone_set('Asia/Taipei');        
  $bdate=date('Y-m-d');  
  $data = $_POST['imageData'];     
  list($type, $data) = explode(';', $data);
  list(, $data)      = explode(',', $data);
  $data = base64_decode($data);  
  file_put_contents("pies/" . $bdate . "_out.jpg", $data);
  
  
  
?>
