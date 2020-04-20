<?php
session_start();
include 'load.php';
$action = isset($_GET['action'])?$_GET['action']:null;
switch ($action) {
    case 'login':
        $ret = new OFFICEconnect();
        $ret->getAuthCode();
        exit;
        break;
        
    default:
        // code...
        break;
}

?>
1. <a href="?action=login" target="_blank" >获取refersh token</a>
<br>
2. <a href="run.php" target="_blank" >运行测试</a>