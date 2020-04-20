<?php
include 'load.php';

header("content-type:application/json");
if(!($refresh_token = file_get_contents(__DIR__ . '/refresh_token')))
    exit('read refresh token failed!!');

$OFFICE = new OFFICEconnect();
$ret = $OFFICE->getRunToken($refresh_token);

if(!empty($ret->refresh_token))
    file_put_contents(__DIR__ . '/refresh_token', $ret->refresh_token);
else {
    echo json_encode($ret) . "\r\n";
    exit('refresh token failed!');
}

$ret = $OFFICE->execAPI($ret->access_token);
file_put_contents('last_run_success', json_encode($ret));

if(!empty($_SERVER['HTTP_REFERER']))
    echo "success\r\n" . json_encode($ret);
else
    echo "success";