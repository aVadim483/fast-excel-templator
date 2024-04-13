<?php
$items = [
    [
        'href' => 'demo1.php',
        'label' => 'demo1.php',
        'img-tpl' => 'img/demo1-tpl.jpg', 'img-out' => 'img/demo1-out.jpg',
    ],
    [
        'href' => 'demo2.php',
        'label' => 'demo2.php',
        'img-tpl' => 'img/demo2-tpl.jpg', 'img-out' => 'img/demo2-out.jpg',
    ],
    [
        'href' => 'demo3.php',
        'label' => 'demo3.php',
        'img-tpl' => 'img/demo3-tpl.jpg', 'img-out' => 'img/demo3-out.jpg',
    ],
];
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Demo FastExcelTemplator</title>
    <style>
        td {vertical-align: top;}
        td img {width: 98%;}
    </style>
</head>
<body>
<table>
    <?php foreach ($items as $n => $item): ?>
    <tr>
        <td colspan="2">
            <h2>Demo <?=$n+1 ?></h2>
            <a href="<?=$item['href'];?>" target="_blank"><?=$item['label'];?></a>
        </td>
    </tr>
    <tr>
        <td><img src="<?=$item['img-tpl'];?>"></td>
        <td>&nbsp;</td>
        <td><img src="<?=$item['img-out'];?>"></td>
    </tr>
    <?php endforeach; ?>
</table>
</body>
</html>