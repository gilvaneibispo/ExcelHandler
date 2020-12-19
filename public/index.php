<?php
require_once '../vendor/autoload.php';

//use app\Init;
use app\ExcelHandler;

$sheet = new ExcelHandler();
$sheet->loadSheet("Matricula_UPT");

$cols = 676;
$lines = 30;

$arr = array();


for ($count = 1; $count <= $cols; $count++) {

    $cc = $sheet->getLetterByIndex($count);

    for ($line = 1; $line <= $lines; $line++) {

        $index = $cc . $line;
        $arr[$index] = "Texto para " . $index;
    }
}

$sheet->setCells($arr);


/*
$sheet->setCells(array(
    "A1" => "A1",
    "A2" => "A2",
    "A3" => "A3",
    "B1" => "B1",
    "B2" => "B2",
    "B3" => "B3",
    "C1" => "C1",
    "C2" => "C2",
    "C3" => "C3"
));
*/
$sheet->saveSheet();

//$init = new Init;
