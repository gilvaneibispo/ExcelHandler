<?php

namespace App;

use App\Handlers\Creator;

class Spreadsheets {

    public function readWorksheet(){

    }

    public function writeWorksheet($config, $header, $data, $fileName = false, $toDownload = false){

        try {

            $st = new Creator($config);
            $st->setHeaders($header);
            $st->insertIntoWorksheet($data);

            $st->saveFile($fileName, $toDownload);
        }catch (\Exception $e){
            throw $e;
        }
    }
}