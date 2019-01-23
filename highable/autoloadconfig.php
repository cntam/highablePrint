<?php
session_start();
error_reporting(0);
if('localhost' ==  $_SERVER['SERVER_NAME']){
    $online = false;  //判断是否在线
    $mac = true;
}elseif ('www.a.cn' ==  $_SERVER['SERVER_NAME']){
    $online = false;  //判断是否在线
    $mac = false;
}elseif ('127.0.0.1' ==  $_SERVER['SERVER_NAME']){
    $online = false;  //判断是否在线
    $mac = false;
}else{
    $online = true;  //判断是否在线
}

if($online){
    require_once '/home/pan/vendor/autoload.php';
}else{
    //require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';  //mac mini
    if($mac){
         require_once '/Users/hongfeitam/vendor/autoload.php';  //macbookPro
    }else{
        require '../../../vendor/autoload.php'; //window10
    }

}
//require '/home/pan/vendor/autoload.php';


define("PRINTURL", "http://allinone321.com/highable/output/"); //打印路徑
define("MSFILEURL", 'http://view.officeapps.live.com/op/view.aspx?src='); //微软网上预览路径