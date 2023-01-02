<?php

namespace App\Handlers;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Creator
{

    protected Spreadsheet $spreadsheet;
    protected Worksheet $sheet;
    protected int $sheetColCounter;
    protected int $sheetLineCounter;
    protected array $headers;
    protected int $startLine;
    protected int $headerLine;
    protected bool $multipleSheets;
    protected bool $toStyleStripes;
    protected string $sheetTitle;

    const TITLE_CELL_BACK_COLOR = "595959";
    const STRIP_CELL_BACK_COLOR = "D9D9D9";
    const CELL_BORDER_COLOR = "898989";
    const CELL_WITH_DEFAULT = 20;
    const CELL_WITH_DOUBLE = 40;
    const CELL_WITH_TRIPLE = 60;

    const INDEX_COL = array(
        'A' => 1, 'B' => 2, 'C' => 3, 'D' => 4, 'E' => 5, 'F' => 6, 'G' => 7, 'H' => 8,
        'I' => 9, 'J' => 10, 'K' => 11, 'L' => 12, 'M' => 13, 'N' => 14, 'O' => 15,
        'P' => 16, 'Q' => 17, 'R' => 18, 'S' => 19, 'T' => 20, 'U' => 21, 'V' => 22,
        'W' => 23, 'X' => 24, 'Y' => 25, 'Z' => 26
    );

    /**
     * Cria uma pasta de trabalho do Excel com uma unica planilha, com o título especificado.
     * @param array $arrConfig - Array de configuração da classe, abaixo as opções e os valores default: <br/>
     * array(<br/>
     *      &emsp;# Linha onde inicia o cabeçalho.<br/>
     *      &emsp;'header_line' => 1,<br/>
     *      &emsp;# Linha onde inicia o corpo da planilha.<br/>
     *      &emsp;'start_line' => 2,<br/>
     *      &emsp;# Criar multiplas planilhas na pasta de trabalho.<br/>
     *      &emsp;'multiple_sheets' => false,<br/>
     *      &emsp;# Adicionar uma coluna com contador de linha.<br/>
     *      &emsp;'show_counter_col' => false,<br/>
     *      &emsp;# Adicionar o estilo com listras em duas cores.<br/>
     *      &emsp;'to_style_stripes' => false,<br/>
     *      &emsp;# Definir um titulo para a planilha única.<br/>
     *      &emsp;'sheet_title' => 'Alunos'<br/>
     * );
     * @throws Exception
     */
    public function __construct(array $arrConfig = array())
    {
        try {

            $this->initAndValidateData($arrConfig);


            # cria uma pasta de trabalho do Excel
            $this->spreadsheet = new Spreadsheet();

            # define a planilha principal como ativa (a que será manipulada)
            $this->spreadsheet->setActiveSheetIndex(0);

            # guarda a instância da planilha ativa
            $this->sheet = $this->spreadsheet->getActiveSheet();

            # define o título da planilha ativa para 'Alunos'
            $this->sheet->setTitle($this->sheetTitle);

            # inicia o array headers
            $this->headers = array();

        } catch (Exception $phpOfficeEx) {
            throw $phpOfficeEx;
        } catch (\Exception $e) {
            throw $e;
        }
    }

    private function initAndValidateData(array $config)
    {

        if (isset($config['header_line'])) {

            if (is_numeric($config['header_line'])) {
                $this->headerLine = $config['header_line'];
            } else {
                throw new \Exception("The header line must be numeric!");
            }
        } else {
            $this->headerLine = 1;
        }

        if (isset($config['start_line'])) {

            if (is_numeric($config['start_line'])) {
                $this->startLine = $config['start_line'];
            } else {
                throw new \Exception("The start line must be numeric!");
            }
        } else {
            $this->startLine = 2;
        }

        if (isset($config['multiple_sheets'])) {

            if (is_bool($config['multiple_sheets'])) {
                $this->multipleSheets = $config['multiple_sheets'];
            } else {
                throw new \Exception("The multiple_sheets flag must be bool!");
            }
        } else {
            $this->multipleSheets = false;
        }

        if (isset($config['show_counter_col'])) {

            if (is_bool($config['show_counter_col'])) {
                $this->showCounterCol = $config['show_counter_col'];
            } else {
                throw new \Exception("The show_counter_col flag must be bool!");
            }
        } else {
            $this->showCounterCol = false;
        }

        if (isset($config['to_style_stripes'])) {

            if (is_bool($config['to_style_stripes'])) {
                $this->toStyleStripes = $config['to_style_stripes'];
            } else {
                throw new \Exception("The to_style_stripes flag must be bool!");
            }
        } else {
            $this->toStyleStripes = false;
        }

        if (isset($config['sheet_title'])) {

            if (is_string($config['sheet_title'])) {
                $this->sheetTitle = $config['sheet_title'];
            } else {
                throw new \Exception("The sheet_title flag must be string!");
            }
        } else {
            $this->sheetTitle = "Planilha_" . date('Y-m-d_H-i-s');
        }

        if($this->headerLine > $this->startLine){
            throw new \Exception("The header_line flag cannot be greater than start_line flag!");
        }

        #$headerLine = isset($config['header_line']) ? $config['header_line'] : 1;
        #$startLine = isset($config['start_line']) ? $config['start_line'] : 2;
        #$multipleSheets = isset($config['multiple_sheets']) ? $config['multiple_sheets'] : false;
        #$showCounterCol = isset($config['start_line']) ? $config['start_line'] : 2;
        #$toStyleStripes = isset($config['start_line']) ? $config['start_line'] : 2;
        #$sheetTitle = isset($config['start_line']) ? $config['start_line'] : 2;
    }

    /**
     * @param array $data - Dados para as linhas da planilha, cada linha é um array.
     * @param array $sheetLabels
     * @return void
     * @throws \Exception
     */
    public function insertIntoWorksheet(array $data, array $sheetLabels = array())
    {
        try {

            if ($this->multipleSheets) {
                $dataFirstLine = (array)$data[0];
                $dataFirstLine = $dataFirstLine[0];
            } else {
                $dataFirstLine = $data[0];
            }

            # converste stdClass para array buscando facilitar/padronizar
            # a contagem de posições...
            if ($dataFirstLine instanceof stdClass) {
                $dataFirstLine = (array)$dataFirstLine;
            }

            # verifica se é um array de arrays analisando o primeiro elemento...
            if (!is_array($dataFirstLine)) {

                throw new \Exception("Data must be an array set!");
            }

            # checa se o numero de colunas das linhas é a mesma que do cabeçalho...
            if ($this->sheetColCounter != count($dataFirstLine)) {

                throw new \Exception("The number of columns does not match the number of header labels!");
            }

            if ($this->multipleSheets)
                $this->insertAndOrganizeInSheets($data, $sheetLabels);
            else
                $this->insertUniqueSheet($data);

        } catch (\Exception $e) {
            throw $e;
        }
    }

    public function setHeaders($headerLabels)
    {
        $this->headers = $headerLabels;
        $this->sheetColCounter = count($headerLabels);
    }

    private function worksheetHeaderCreate()
    {

        try {
            $countCol = 1;

            if (count($this->headers) != 0) {

                foreach ($this->headers as $header) {

                    $this->setCellDataValue($countCol, $this->headerLine, $header);

                    $countCol++;
                }
            } else {

                throw new \Exception("Header array cannot be empty!");
            }
        } catch (\Exception $e) {
            throw $e;
        }
    }

    public function setColFormat(array $formatByCol)
    {

        foreach ($formatByCol as $col => $format) {

            $endLine = ($this->sheetLineCounter + ($this->startLine - 1));

            $this->sheet
                ->getStyle("{$col}{$this->startLine}:{$col}{$endLine}")
                ->getNumberFormat()
                ->setFormatCode($format);
        }
    }

    public function setColWidth(array $indexAndWidth)
    {

        try {
            foreach ($indexAndWidth as $letter => $width) {
                #echo "{$index} => {$width}<br/>";
                $index = self::getIndexByColLatter($letter);
                $this->sheet->getColumnDimensionByColumn($index)->setWidth($width);
                $this->sheet->getColumnDimensionByColumn($index)->setAutoSize(false);
            }
        } catch (\Exception $e) {
            throw $e;
        }
    }

    public function saveFile($fileName = null, $toDownload = false, $filePath = "")
    {

        try {
            //$writer = new Xlsx($this->spreadsheet);
            $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');

            $fileName = ($fileName == null)?$this->sheetTitle : $fileName;

            if ($toDownload) {

                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="' . $fileName . '.xlsx"');
                header('Cache-Control: max-age=0');

                $writer->save('php://output');
            } else {

                $fileLocation = dirname(__DIR__, 2);
                $filePath = ($filePath == "" ? "/{$fileName}" : "/{$filePath}/{$fileName}");
                $writer->save("{$fileLocation}{$filePath}.xlsx");
            }
        } catch (\Exception $e) {
            throw $e;
        }
    }

    public function worksheetHeaderStylize()
    {

    }

    /**
     * Aplica uma borda em volta o documento
     * @throws \Exception
     */
    public function applyDocBorder()
    {

        try {
            $styleArray = array(
                'borders' => array(
                    'outline' => array(
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => array('argb' => self::CELL_BORDER_COLOR),
                    ),
                ),
            );

            $lastLine = $this->sheetLineCounter + $this->startLine - 1;
            $endSheet = self::getLetterByColIndex($this->sheetColCounter) . $lastLine;
            $this->sheet->getStyle("A{$this->headerLine}:{$endSheet}")->applyFromArray($styleArray);
        } catch (\Exception $e) {
            throw $e;
        }
    }

    public function applyStyleStripes()
    {

        $endSheetCol = $this->applyHeaderStyle();

        for ($line = $this->startLine; $line < ($this->sheetLineCounter + $this->startLine - 1); $line++) {

            $cellLine = $this->sheet->getStyle("A{$line}:{$endSheetCol}{$line}");

            $bkCell = $cellLine->getFill()->setFillType(Fill::FILL_SOLID);

            if ($line % 2 == 0) {
                $bkCell->getStartColor()->setARGB(self::STRIP_CELL_BACK_COLOR);
            } else {
                $bkCell->getStartColor()->setARGB(Color::COLOR_WHITE);
            }
        }

    }

    private function applyHeaderStyle(): string
    {

        try {
            $endSheetCol = self::getLetterByColIndex($this->sheetColCounter);
            $lineTitle = $this->sheet->getStyle("A{$this->headerLine}:{$endSheetCol}{$this->headerLine}");
            $lineTitle->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            $lineTitle->getFont()->setBold(true);
            $bkCell = $lineTitle->getFill()->setFillType(Fill::FILL_SOLID);
            $bkCell->getStartColor()->setARGB(self::TITLE_CELL_BACK_COLOR);

            return $endSheetCol;
        }catch (\Exception $e){
            throw $e;
        }
    }

    private function insertUniqueSheet($data)
    {
        try {
            # cria o header da planilha
            $this->worksheetHeaderCreate();

            $this->sheetLineCounter = count($data);
            $line = $this->startLine;

            # percorre linha por linha
            foreach ($data as $dataLine) {

                # convertendo stdClass para array
                if ($dataLine instanceof stdClass) {
                    $dataLine = (array)$dataLine;
                }

                # o número da coluna é sempre 1 no inicio da linha
                $countCol = 1;

                # percorre coluna por coluna na linha atual
                foreach ($dataLine as $dataCell) {

                    # define o valor para a célula atual
                    $this->setCellDataValue($countCol, $line, $dataCell);
                    $this->setColWidth2($countCol);
                    $countCol++;
                }

                $line++;
            }

            if ($this->toStyleStripes)
                $this->applyStyleStripes();
            $this->applyDocBorder();

        } catch (\Exception $e) {
            throw $e;
        }
    }

    private function insertAndOrganizeInSheets(array $data, array $sheetTitles)
    {

        try {
            for ($i = 0; $i < count($data); $i++) {

                if ($i != 0) {
                    $this->createAndActiveWorksheet($i, $sheetTitles[$i]);
                } else {
                    $this->sheet->setTitle($sheetTitles[0]);
                }

                $this->insertUniqueSheet($data[$i]);
            }
        } catch (\Exception $e) {
            throw $e;
        }
    }

    private function setCellDataValue($colNumber, $lineNumber, $data)
    {

        try {
            # recupera a letra equivalente ao numero da coluna atual
            $letter = $this->getLetterByColIndex($colNumber);

            # forma o index da célula, no formato coluna (letra) e linha, ex.: 'A1'
            $cell = "{$letter}{$lineNumber}";

            # define/atualiza o valor para a célula do título
            $this->sheet->setCellValue($cell, $data);

        } catch (\Exception $e) {
            throw $e;
        }
    }

    private function setColWidth2($colNumber, $width = 30){

        try {
            $this->sheet->getColumnDimensionByColumn($colNumber)->setWidth($width);
            $this->sheet->getColumnDimensionByColumn($colNumber)->setAutoSize(false);
        } catch (\Exception $e) {
            throw $e;
        }
    }

    private function createAndActiveWorksheet(int $index, string $title)
    {

        try {

            $this->spreadsheet->createSheet($index);
            $this->spreadsheet->setActiveSheetIndex($index);
            $this->sheet = $this->spreadsheet->getActiveSheet();
            $this->sheet->setTitle($title);
        } catch (\Exception $e) {
            throw $e;
        }
    }

    private static function getIndexByColLatter(string $letters): int
    {

        $sizeLetters = strlen($letters);

        if ($sizeLetters == 0) {
            # a string esta vazia, informar valor válido...
            throw new \Exception("A string esta vazia, informar valor válido!");
        } elseif ($sizeLetters == 1) {
            return self::INDEX_COL[$letters];
        } elseif ($sizeLetters == 2) {
            $l1 = \app\mb_strtoupper($letters[0]);
            $l2 = \app\mb_strtoupper($letters[1]);

            $i1 = self::INDEX_COL[$l1];
            $i2 = self::INDEX_COL[$l2];

            return ($i1 * 26) + $i2;
        } else {
            throw new \Exception("Apenas duas letras são suportadas!");
        }
    }

    private static function getLetterByColIndex($index): string
    {


        # o index deve está entre 1 e 702;
        # 702 vem do fato que são aceitas duas letras onde o maximo é ZZ
        # A primerira letra Z representa 26*26 = 676.
        # a segunda letra Z representa mais 26...
        # logo o maximo é $index = (26*26)+26 = 702;
        if ($index > 0 && $index <= 702) {

            # Encontra um divisor flutuante, ex.: 58/26=2.23
            $i1 = ($index / 26);

            # Pega a parte iteira do divisor encontrado
            $i2 = (int)$i1;

            # define a diferença
            $i3 = $index - ($i2 * 26);

            if ($i3 == 0) {
                $i2 = $i2 - 1;
                $i3 = 26;
            }

            return self::singlePosLetter($i2) . self::singlePosLetter($i3);
        } else {

            throw new \Exception('The index must be between 1 and 675!');
        }
    }

    private static function singlePosLetter(string $val)
    {
        return array_search($val, self::INDEX_COL);
    }
}