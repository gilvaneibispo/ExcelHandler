<?php
require_once '../vendor/autoload.php';

use app\ExcelHandler;

$sheet = new ExcelHandler();
$sheet->loadSheet("Matricula_UPT");

$numberTabs = $sheet->getTabCount();
$tabs = $sheet->getTabNames();
$latter = $sheet->getLetterCol();
$cols = $sheet->getIndexCol($latter);
$textFilters = array(
    "Inscritos",
    "Não inscritos"
);


foreach ($tabs as $tab) {

    $sheet->getTabByName($tab);
    $rows = $sheet->getNumRows();
    $realSize = 0;

    echo "Criando planilha para {$tab} com {$rows} linhas! <br/>";

    $dataNewSheet = array();

    for ($line = 1; $line <= $rows; $line++) {

        if(isValidCell($sheet->getDataCell("A{$line}"), $textFilters)){
            $realSize++;
        }

        for ($col = 1; $col <= $cols; $col++) {

            $currentCol = $sheet->getLetterByIndex($col);

            $theVal = $sheet->getDataCell("{$currentCol}{$line}");
            $theVal = trim($theVal);

            if (isValidCell($theVal, $textFilters)) {

                /* title line */
                if ($line == 1) {

                    $dataNewSheet['A1'] = "#";

                    if ($currentCol == 'C') {
                        $dataNewSheet['B1'] = $theVal;
                    } elseif ($currentCol == 'S') {
                        $dataNewSheet['C1'] = "WhatsApp";
                    } elseif ($currentCol == "U") {
                        $dataNewSheet['D1'] = $theVal;
                    } elseif ($currentCol == "AG") {
                        $dataNewSheet['E1'] = $theVal;
                        $dataNewSheet['F1'] = "Frequência do mês (dez/20)";
                    }
                } elseif ($line == 2) {
                    if ($currentCol == 'AG') {

                        // From cell F (index 6) at cell AJ (index 36)...
                        for($x = 6;$x <= 36; $x++){
                            $let = $sheet->getLetterByIndex($x);
                            $dataNewSheet["{$let}2"] = ($x - 5);
                        }
                    }
                } else {
                    $dataNewSheet['A' . $line] = ($line - 2);
                    if ($currentCol == 'C') {
                        $dataNewSheet['B' . $line] = mb_strtoupper($theVal, 'UTF-8');
                    } elseif ($currentCol == 'S') {
                        $dataNewSheet['C' . $line] = $theVal;
                    } elseif ($currentCol == "U") {
                        $dataNewSheet['D' . $line] = strtolower($theVal);
                    } elseif ($currentCol == "AG") {
                        $dataNewSheet['E' . $line] = $theVal;
                    }
                }
            }
        }
    }

    $tabName = mb_strtoupper($tab, 'UTF-8');
    $tabName = "spaces/{$tabName}.xlsx";

    $sheet->sheetCreate($tabName, $dataNewSheet, $realSize);
}

function isValidCell($val, $arr)
{
    return ((in_array($val, $arr) || empty($val)) ? false : true);
}