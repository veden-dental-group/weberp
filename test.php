<?php
date_default_timezone_set('Asia/Taipei');
// $tdate=date('Y-m-d');

$tdate = $argv[1] ? $argv[1] : date('Y-m-d');
echo $edate = date("Y-m-d", strtotime($tdate) - (86400 * 4));
?>
