<?php
session_start();
header("content-type:application/json");
include __DIR__ . '/load.php';

$ret = new OFFICEconnect();
$info = $ret->getAccessToken();

if(isset($info->refresh_token))
    echo "获取 refresh_token 成功\r\n";
else
{
    echo '获取 refresh_token 失败';
    exit;
}
file_put_contents('refresh_token', '');
file_put_contents('refresh_token', $info->refresh_token);

if(!empty(file_get_contents('refresh_token')))
    echo "refresh_token 存储成功\r\n请添加[http(s)://域名/run.php]到Corn任务";
else
    echo 'refresh_token 存储失败';