<?php
if('localhost' ==  $_SERVER['SERVER_NAME']){
    $online = false;  //判断是否在线
}else{
    $online = true;  //判断是否在线
}

if($online){
    require_once '/home/pan/vendor/autoload.php';

}else{
    require_once '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
}
//require '/home/pan/vendor/autoload.php';

//require '../vendor/autoload.php';


//if($online){
//    echo '/home/pan/vendor/autoload.php';
//
//}else{
//    echo '/Applications/XAMPP/xamppfiles/htdocs/composer/vendor/autoload.php';
//}