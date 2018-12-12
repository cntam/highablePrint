<?php
if('localhost' ==  $_SERVER['SERVER_NAME']){
    $online = false;  //判断是否在线
}elseif ('www.a.cn' ==  $_SERVER['SERVER_NAME']){
    $online = false;  //判断是否在线
}elseif ('127.0.0.1' ==  $_SERVER['SERVER_NAME']){
    $online = false;  //判断是否在线
}else{
    $online = true;  //判断是否在线
}

if($online){
    require_once '/home/pan/vendor/autoload.php';
}else{
    //require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';  //mac mini
    require_once '/Users/hongfeitam/vendor/autoload.php';  //macbookPro
    //require '../vendor/autoload.php';  //window10
}
//require '/home/pan/vendor/autoload.php';


define("PRINTURL", "http://allinone321.com/highable/output/"); //打印路徑
define("MSFILEURL", 'http://view.officeapps.live.com/op/view.aspx?src='); //微软网上预览路径

//if($online){
//    echo '/home/pan/vendor/autoload.php';
//
//}else{
//    echo '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
//}