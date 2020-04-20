<?php
// +----------------------------------------------------------------------
// | Copyright (c) 2020.
// +----------------------------------------------------------------------
// | Author: 祭夜 <me@jysafe.cn>
// +----------------------------------------------------------------------
/**
 *  OFFICE第三方登录认证
 */
class OFFICEconnect
{
    private static $data;
    //APP ID
    private $app_id = "";
    //APP KEY
    private $app_key = "";
    //回调地址
    private $callBackUrl = "";
    //Authorization Code
    private $code = "";
    //access Token
    private $accessToken = "";
    //role
    private $role = "";

    public function __construct()
    {
        $this->app_id = CLIENT_ID; //App Key
        $this->app_key = CLIENT_PASS; //App Secret
        $this->callBackUrl = 'https://' . OFFICE_CALLBACK_DOMAIN . '/office.callback.php';
        $this->code = isset($_GET['code'])?$_GET['code']:null;

        //检查用户数据
        if (empty($_SESSION['QC_userData'])) {
            self::$data = array();
        } else {
            self::$data = $_SESSION['QC_userData'];
        }
    }

    //获取Authorization Code
    public function getAuthCode()
    {
        $url = "https://login.microsoftonline.com/" . OFFICE_TENANT . "/oauth2/v2.0/authorize";
        $param['client_id'] = $this->app_id;
        $param['response_type'] = "code";
        $param['redirect_uri'] = $this->callBackUrl;
        // $param['response_mode'] = 'query';
        $param['scope'] = OFFICE_SCOPE;
        $param['access_type'] = 'offline';
        //-------生成唯一随机串防CSRF攻击
        $state = md5(uniqid(rand(), TRUE));
        $param['state'] = $state;
        $param = http_build_query($param, '', '&');
        $url = $url . "?" . $param;
        $_SESSION['state'] = $state;
        // exit($url);
        header("Location:" . $url);
    }

    //通过Authorization Code获取Access Token
    public function getAccessToken()
    {
        $state = isset($_GET['state'])?$_GET['state']:exit('can`t found state');
        if($state !== $_SESSION['state'])
        {
            echo $_SESSION['state'];
            exit('state error');
        }
        $url = "https://login.microsoftonline.com/" . OFFICE_TENANT . "/oauth2/v2.0/token";
        $param['client_id'] = $this->app_id;
        $param['scope'] = OFFICE_SCOPE;
        $param['code'] = $this->code;
        $param['redirect_uri'] = $this->callBackUrl;
        $param['grant_type'] = "authorization_code";
        $param['client_secret'] = $this->app_key;
        $param = http_build_query($param, '', '&');
        $data = $this->postUrl($url, $param);
        $data = json_decode($data);
        $_SESSION['access_token'] = $data->access_token;
        return $data;
    }

    //获取执行token
    public function getRunToken($refresh_token)
    {
        $url = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        $param['client_id'] = $this->app_id;
        $param['grant_type'] = "refresh_token";
        $param['scope'] = OFFICE_SCOPE;
        $param['refresh_token'] = $refresh_token;
        $param['client_secret'] = $this->app_key;
        $param['redirect_uri'] = $this->callBackUrl;
        
        $param = http_build_query($param, '', '&');
        // echo $param . "\r\n";
        //$url=$url."?".$param;
        $data = $this->postUrl($url, $param);
        $data = json_decode($data);
        $_SESSION['access_token'] = $data->access_token;
        // var_dump($data);
        return $data;
    }
    
    // 执行API
    public function execAPI($access_token)
    {
        $url[] = 'https://graph.microsoft.com/beta/administrativeUnits';
        
        $url[] = 'https://graph.microsoft.com/v1.0/me/drive/root';
        $url[] = 'https://graph.microsoft.com/v1.0/me/drive';
        $url[] = 'https://graph.microsoft.com/v1.0/users';
        $url[] = 'https://graph.microsoft.com/v1.0/me';
        $url[] = 'https://graph.microsoft.com/v1.0/me/messages';
        $url[] = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules';
        $url[] = 'https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages';
        $url[] = 'https://graph.microsoft.com/v1.0/me/drive/root/children';
        $url[] = 'https://api.powerbi.com/v1.0/myorg/apps';
        $url[] = 'https://graph.microsoft.com/v1.0/me/mailFolders';
        $url[] = 'https://graph.microsoft.com/v1.0/me/outlook/masterCategories';
        
        $headers[] = 'Authorization:Bearer ' . $access_token;
        $headers[] = 'Content-Type: application/json';
        
        foreach ($url as $value) {
            // echo "\r\n--->" . $value . "\r\n";
            $data = $this->getUrl($value, $headers);
            $ret[] = $data;
        }
        return $ret;
    }
    
    //CURL GET
    private function getUrl($url, $headers = null)
    {
        $ch = curl_init($url);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($ch, CURLOPT_TIMEOUT, 5);
        curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
        // if (!empty($options)){
        //     curl_setopt_array($ch, $options);
        // }
        $data = curl_exec($ch);
        curl_close($ch);
        return $data;
    }

    //CURL POST
    private function postUrl($url, $data)
    {
        $ch = curl_init();
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, TRUE);
        curl_setopt($ch, CURLOPT_POST, TRUE);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $data);
        curl_setopt($ch, CURLOPT_URL, $url);
        curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, FALSE);
        $ret = curl_exec($ch);
        curl_close($ch);
        return $ret;
    }
}
