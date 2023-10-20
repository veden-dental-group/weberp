<?php
    include_once("_data.php"); 
    session_start();
                    
    $_SESSION['name']       = "";  
    $_SESSION['account']    = "";   
    $_SESSION['userguid']   = "";  
    $_SESSION['maker']      = ""; 
    $_SESSION['lastlogin']  = "";
    forward("index.php");
                                  
?>
