<?php

//autoload do projeto
require '../vendor/autoload.php';

//classe responsável pela manipulação da planilha
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//classe que salvará a planilha em .xlsx
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//classe responsável pelo load dos arquivos de planilha
use PhpOffice\PhpSpreadsheet\IOFactory;

getIndexCol("AADX");

function loadSheet()
{

    $arrPosition = ['A', 'B', 'C', 'D', 'E', 'F'];
    $hasCol = true;
    $hasLine = true;
    $latterPos = 0;
    $line = 1;

    //carregando a planilha spreadsheet2 em um objeto PHP
    $spreadsheet = IOFactory::load('spreadsheet1.xlsx');
    //retornando a aba ativa
    $sheet = $spreadsheet->getActiveSheet();

   
    

    //var_dump($sheet->getHighestRowAndColumn());

    /*while($hasLine){

        while($hasCol){
            
            $position = $arrPosition[$latterPos] . $line;
            $textValue = $sheet->getCell($position)->getValue();

            if(empty($textValue)){
                $hasCol = false;
                $textValue = "endline";
            }

            $latterPos++;
            echo "{$textValue} | ";
        }

        if(empty($textValue)){
            $hasCol = false;
            $hasLine = false;
        }else{
            $hasCol = true;
        }

        $line++;
    }*/

    //$cellA1 recebe os dados da célula A1
    $cellA1 = $sheet->getCell('A1')->getValue();
    //$cellD1 recebe os dados da célula D1
    $cellD1 = $sheet->getCell('D1')->getValue();
    //$cellB2 recebe os dados da célula B2
    $cellB2 = $sheet->getCell('B2')->getValue();

    //retornando os dados da célula A1
    echo ('A1 = ' . $cellA1 . PHP_EOL . "<br/>");
    //retornando os dados da célula D1
    echo ('D1 = ' . $cellD1 . PHP_EOL . "<br/>");
    //retornando os dados da célula B2
    echo ('B2 = ' . $cellB2 . PHP_EOL . "<br/>");
}

function saveSheet()
{

    //instanciando uma nova planilha
    $spreadsheet = new Spreadsheet();

    //retornando a aba ativa
    $sheet = $spreadsheet->getActiveSheet();

    $sheet = sheetCreate($sheet);

    //Instanciando uma nova planilha
    $writer = new Xlsx($spreadsheet);
    //salvando a planilha na extensão definida
    $writer->save('spreadsheet1.xlsx');

    echo "Planilha salva em disco...";
}

function sheetCreate($sheet)
{

    //Definindo a célula A1
    $sheet->setCellValue('A1', 'Nome');

    //Definindo a célula B1
    $sheet->setCellValue('B1', 'Nota 1');

    $sheet->setCellValue('C1', 'Nota 2');

    $sheet->setCellValue('D1', 'Media');

    $sheet->setCellValue('A2', 'pokemaobr');

    $sheet->setCellValue('B2', 5);

    $sheet->setCellValue('C2', 3.5);

    //Definindo a fórmula para o cálculo da média
    $sheet->setCellValue('D2', '=((B2+C2)/2)');

    $sheet->setCellValue('A3', 'bob');

    $sheet->setCellValue('B3', 7);

    $sheet->setCellValue('C3', 8);

    $sheet->setCellValue('D3', '=((B3+C3)/2)');

    $sheet->setCellValue('A4', 'boina');

    $sheet->setCellValue('B4', 9);

    $sheet->setCellValue('C4', 9);

    $sheet->setCellValue('D4', '=((B4+C4)/2)');

    return $sheet;
}

function getIndexCol($col)
{

    $columnLookup = [
        'A' => 1, 'B' => 2, 'C' => 3, 'D' => 4, 'E' => 5, 'F' => 6, 'G' => 7, 'H' => 8,
        'I' => 9, 'J' => 10, 'K' => 11, 'L' => 12, 'M' => 13, 'N' => 14, 'O' => 15,
        'P' => 16, 'Q' => 17, 'R' => 18, 'S' => 19, 'T' => 20, 'U' => 21, 'V' => 22,
        'W' => 23, 'X' => 24, 'Y' => 25, 'Z' => 26
    ];

    $len = strlen($col);
    $maxLen = strlen($col);
    $minLen = ($len - 1) * 26;
    
    $last = substr($col, $len - 1, 1);

    echo "Pos " . ($columnLookup[$last] + $minLen);
}
