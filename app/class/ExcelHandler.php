<?php


namespace app;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ExcelHandler
{
    private $sheet;
    private $spreadsheet;
    //private $countCols;
    private $refIndexCol;

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();

        $this->refIndexCol = array(
            'A' => 1, 'B' => 2, 'C' => 3, 'D' => 4, 'E' => 5, 'F' => 6, 'G' => 7, 'H' => 8,
            'I' => 9, 'J' => 10, 'K' => 11, 'L' => 12, 'M' => 13, 'N' => 14, 'O' => 15,
            'P' => 16, 'Q' => 17, 'R' => 18, 'S' => 19, 'T' => 20, 'U' => 21, 'V' => 22,
            'W' => 23, 'X' => 24, 'Y' => 25, 'Z' => 26
        );

        //$sheetIndex = $this->spreadsheet->getIndex(
        //    $this->spreadsheet->getSheetByName('Geral')
        //);
        $this->spreadsheet->removeSheetByIndex(0);

        //$myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($this->spreadsheet, 'My Data');

        // Attach the "My Data" worksheet as the first worksheet in the Spreadsheet object
        //$this->spreadsheet->addSheet($myWorkSheet);
    }

    public function getTabByIndex($index)
    {
        $this->sheet = $this->spreadsheet->getSheet($index);
    }

    public function getTabByName($tabName){
        $this->sheet = $this->spreadsheet->getSheetByName($tabName);
    }

    public function getTabCount()
    {
        return $this->spreadsheet->getSheetCount();
    }

    public function getTabNames()
    {
        return $this->spreadsheet->getSheetNames();
    }

    public function sheetCreate($name, $cells, $rows)
    {
        $TITLE_CELL_BACK_COLOR = "595959";
        $STRIP_CELL_BACK_COLOR = "D9D9D9";
        $spread = new Spreadsheet();

        //$myWorkSheet = new Worksheet($spread, 'Alunos');

        //$spread->addSheet($myWorkSheet);
        $sheet = $spread->getActiveSheet();
        $sheet->setTitle("Alunos");

        for($r = 3; $r <= $rows; $r++) {
            for ($c = 6; $c <= 36; $c++) {

                $l2 = $this->getLetterByIndex($c);
                $cellLine = $sheet->getStyle("{$l2}{$r}");
                $bkCell = $cellLine->getFill()->setFillType(Fill::FILL_SOLID);

                if ($r % 2 == 0) {
                    $bkCell->getStartColor()->setARGB($STRIP_CELL_BACK_COLOR);
                }else{
                    $bkCell->getStartColor()->setARGB(Color::COLOR_WHITE);
                }

                $cellLine->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            }
        }

        //$idCell = 1;
        foreach ($cells as $cellKey => $cellValue) {

            $sheet->setCellValue($cellKey, $cellValue);
            $cellLine = $sheet->getStyle($cellKey);
            $cellLine->getNumberFormat()->setFormatCode("#");
            $cellLine->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

            $idCell = preg_replace('/[^0-9]/', '', $cellKey);

            $bkCell = $cellLine->getFill()->setFillType(Fill::FILL_SOLID);
            if($idCell % 2 == 0){
                $bkCell->getStartColor()->setARGB($STRIP_CELL_BACK_COLOR);
            }else{
                $bkCell->getStartColor()->setARGB(Color::COLOR_WHITE);
            }

            if(substr($cellKey, 0, 1) == "A"){
                $cellLine->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
            }

            if($cellKey == "A1"){
                $cellLine->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            }

            //$idCell++;
        }

        for($i = 1; $i <= 36; $i++){

            $l = $this->getLetterByIndex($i);

            $lineTitle = $sheet->getStyle("{$l}1");
            $lineTitle->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            $lineTitle->getFont()->setBold(true);
            $bkCell = $lineTitle->getFill()->setFillType(Fill::FILL_SOLID);
            $bkCell->getStartColor()->setARGB($TITLE_CELL_BACK_COLOR);

            $lineSubtitle = $sheet->getStyle("{$l}2");
            $lineSubtitle->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            $lineSubtitle->getFont()->setBold(true);
            $bkCell = $lineSubtitle->getFill()->setFillType(Fill::FILL_SOLID);
            $bkCell->getStartColor()->setARGB($TITLE_CELL_BACK_COLOR);
            $lineSubtitle->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }

        for($c = 1; $c <= 36; $c++){

            $l2 = $this->getLetterByIndex($c);
            $width = "5";

            if($l2 == "B"){
                $width = "42";
            }elseif ($l2 == "C"){
                $width = "17";
            }elseif ($l2 == "D"){
                $width = "36";
            }elseif ($c == 5){
                $width = "17";
            }elseif ($l2 = "E"){
                $width = "4";
            }

            $sheet->getColumnDimensionByColumn($c)->setWidth($width);
            $sheet->getColumnDimensionByColumn($c)->setAutoSize(false);
        }

        $sheet->mergeCells("F1:AJ1");
        $sheet->getStyle("F1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->mergeCells("A1:A2");
        $sheet->getStyle("A1")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle("A1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->mergeCells("B1:B2");
        $sheet->getStyle("B1")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle("B1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->mergeCells("C1:C2");
        $sheet->getStyle("C1")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle("C1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->mergeCells("D1:D2");
        $sheet->getStyle("D1")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle("D1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $sheet->mergeCells("E1:E2");
        $sheet->getStyle("E1")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle("E1")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $writer = new Xlsx($spread);
        $writer->save($name);
    }

    public function setCells($cells = array())
    {

        foreach ($cells as $cellKey => $cellValue) {
            $this->sheet->setCellValue($cellKey, $cellValue);
        }
    }

    public function setCell($cellKey, $cellValue)
    {

        $this->sheet->setCellValue($cellKey, $cellValue);
    }

    public function getDataCells($keys = array())
    {

        $allCells = array();

        foreach ($keys as $key) {
            $allCells[$key] = $this->sheet->getCell($key)->getValue();
        }

        if (count($keys) == 0) {

        }

        return $allCells;
    }

    public function getDataCell($key)
    {

        return $this->sheet->getCell($key)->getValue();
    }

    function saveSheet($fileName = null)
    {
        try {

            $name = date('Ymd_Hi');
            $name = "spreadsheet_" . $name;
            $name = ($fileName != null ? $fileName : $name);
            $name = $name . '.xlsx';

            //$sheet = $this->spreadsheet->getActiveSheet();
            //$sheet = sheetCreate($sheet);
            $writer = new Xlsx($this->spreadsheet);

            $writer->save($name);

            echo "Saved!<br/>";
            return true;
        } catch (\Exception $e) {
            return false;
        }
    }

    public function loadSheet($sheetName)
    {
        $this->spreadsheet = IOFactory::load("{$sheetName}.xlsx");
        $this->sheet = $this->spreadsheet->getActiveSheet();

        $this->spreadsheet->removeSheetByIndex(0);

        $myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($this->spreadsheet, 'My Data');

        // Attach the "My Data" worksheet as the first worksheet in the Spreadsheet object
        $this->spreadsheet->addSheet($myWorkSheet);

        echo "Loading...<br/>";
    }

    public function getNumCols()
    {

        //return $this->getDimension()[''];
    }

    public function getNumRows()
    {
        return $this->getDimension('row');
    }

    public function getLetterCol()
    {
        return $this->getDimension('col');
    }

    private function getDimension($dir = null)
    {

        $dir = ($dir == "col" ? "column" : $dir);
        $letterAndNumber = $this->sheet->getHighestRowAndColumn();

        if (array_key_exists($dir, $letterAndNumber)) {
            return $letterAndNumber[$dir];
        } else {
            return array(
                'rows' => $letterAndNumber['row'],
                'letter' => $letterAndNumber['column']
            );
        }
    }


    public function getIndexCol($letterCol)
    {
        $len = strlen($letterCol);

        if ($len == 1) {
            $numCols = $this->refIndexCol[$letterCol];
        } elseif ($len == 2) {

            // i1 = index first letter; i2 = index last letter.
            // p = (26 * i1) + i2; Ex.: 'BC' => p = (26*2) + 3 = 55;
            $letters = str_split($letterCol);
            $i1 = $letters[0];
            $i2 = $letters[1];
            //echo "{$i1} {$i2}" . $this->singlePosLetter('B') . " - " . $this->singlePosLetter('F');
            $numCols = (26 * $this->refIndexCol[$i1]) + $this->refIndexCol[$i2];
        } else {
            throw new \Exception('The columns number cannot exceed 676 [3 letters in string]!');
        }

        return $numCols;
    }

    public function getLetterByIndex($index)
    {

        if ($index == 0 || $index > 676) {
            throw new \Exception('The index must be between 1 and 676!');
        } else {
            $i1 = ($index / 26);
            $i2 = (int)$i1;
            $i3 = $index - ($i2 * 26);

            if ($i3 == 0) {
                $i2 = $i2 - 1;
                $i3 = 26;
            }

            return $this->singlePosLetter($i2) . $this->singlePosLetter($i3);
        }
    }

    public function singlePosLetter(string $val)
    {
        return array_search($val, $this->refIndexCol);
    }
}