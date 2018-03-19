<?php
//模式分隔符后的"i"标记这是一个大小写不敏感的搜索
if (preg_match("/php/i", "PHP is the web scripting language of choice.")) {
    echo "查找到匹配的字符串 php。";
} else {
    echo "未发现匹配的字符串 php。";
}

echo "-------------------------------";

// 从URL中获取主机名称
preg_match('@^(?:http://)?([^/]+)@i',
    "http://www.runoob.com/index.html", $matches);
$host = $matches[1];
echo $host;
var_dump($matches);
echo "<br>-------------------------------";
echo "<br>";

$img = 'http://www.a.cn/wordpress/wp-content/uploads/2018/03/%E9%A6%96%E9%A0%81_20171222113533.png';
preg_match ('/.(jpg|gif|bmp|jpeg|png)/i', $img, $imgformat);
$imgformat = $imgformat[1];
switch ($imgformat)
{
    case "jpg":
    case "jpeg":
        $img = imagecreatefromjpeg($img);
        break;
    case "bmp":
        $img =  imagecreatefromwbmp($img);
        break;
    case "gif":
        $img =  imagecreatefromgif($img);
        break;
    case "png":
        $img =   imagecreatefrompng($img);
        break;
}
echo "<br>";
echo $imgformat;
//
echo "<br>";
$width = imagesx($img);
echo $width;
?>