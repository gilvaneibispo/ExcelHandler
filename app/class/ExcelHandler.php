<?php


namespace app;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;


class ExcelHandler
{
    private $sheet;
    private $spreadsheet;
    //private $countCols;
    private $refIndexCol = array(
        'A' => 1, 'B' => 2, 'C' => 3, 'D' => 4, 'E' => 5, 'F' => 6, 'G' => 7, 'H' => 8,
        'I' => 9, 'J' => 10, 'K' => 11, 'L' => 12, 'M' => 13, 'N' => 14, 'O' => 15,
        'P' => 16, 'Q' => 17, 'R' => 18, 'S' => 19, 'T' => 20, 'U' => 21, 'V' => 22,
        'W' => 23, 'X' => 24, 'Y' => 25, 'Z' => 26
    );

    public function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    public function sheetCreate()
    {

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

    public function loadSheet()
    {
        $this->spreadsheet = IOFactory::load('spreadsheet1.xlsx');
        $this->sheet = $this->spreadsheet->getActiveSheet();

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
            $numCols = (26 * $i1) + $i2;
        } else {
            throw new \Exception('The columns number cannot exceed 676 [3 letters in string]!');
        }

        return $numCols;
    }

    public function getLetterByIndex($index)
    {

        if($index == 0 || $index > 676){
            throw new \Exception('The index must be between 1 and 676!');
        }else {
            $i1 = ($index / 26);
            $i2 = (int) $i1;
            $i3 = $index - ($i2 * 26);

            if($i3 == 0){
                $i2 = $i2 - 1;
                $i3 = 26;
            }

            return $this->singlePosLetter($i2) . $this->singlePosLetter($i3);
        }
    }

    private function singlePosLetter(string  $val){
        return array_search($val, $this->refIndexCol);
    }
}