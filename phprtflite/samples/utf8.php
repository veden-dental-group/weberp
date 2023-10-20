<?php

$dir = dirname(__FILE__);
require_once $dir . '/../lib/PHPRtfLite.php';

// register PHPRtfLite class loader
PHPRtfLite::registerAutoloader();

//Rtf document
$rtf = new PHPRtfLite();

//Font
$times12 = new PHPRtfLite_Font(12, 'Times new Roman');

//Section
$sect = $rtf->addSection();
//Write utf-8 encoded text.
//Text is from file. But you can use another resouce: db, sockets and other
//$sect->writeText(file_get_contents($dir . '/sources/utf8.txt'), $times12, null);
$font1 = new PHPRtfLite_Font('12', 'Tahoma', '#00ff00'); 
$font2 = new PHPRtfLite_Font('12', 'Tahoma', '#ff0000'); 
$sect->writeText('到貨數:', $font1);
$sect->writeText('133', $font2);     
$sect->writeText('顆  ', $font1);
$sect->writeText('內返數:', $font1);
$sect->writeText('55', $font2);     
$sect->writeText('顆  ', $font1);
$sect->writeText('返修:', $font1);
$sect->writeText('99', $font2);     
$sect->writeText('顆  ', $font1);


// save rft document
$rtf->save($dir . '/generated/' . basename(__FILE__, '.php') . '.rtf');
