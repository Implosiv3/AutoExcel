<?php

require_once '../myClasses/MyExcel.php';
require_once '../myClasses/MyExcelFormatter.php';

if (isset($_POST["submit"])) {
    $file = 'C:\xampp\htdocs\MyExcel_Pro\files\ejemplo.xlsx';
    $table = MyExcel::readTable($file, 0, 'E', 4, 'G', 9);
    $headers = $table[0];
    array_splice($table, 0, 1);
    echo MyExcelFormatter::tableToHTML_index($headers, $table);

    MyExcel::getContent($file, 0);
} else {
    echo '...';
}