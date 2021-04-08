<?php
namespace App\Repositories;

use Exception;
use \PhpOffice\PhpSpreadsheet\IOFactory;
use \PhpOffice\PhpSpreadsheet\Cell\Coordinate;

/**
 * Processing Excel file class
 */
class EngineFiles {

    protected $reader;

    protected $spreadsheet;

    protected $tables;

    static $fileExts = ['xls', 'xlsx'];

    public function __construct($filePath)
    {
        $arrPath = @explode(".", $filePath);

        if( ! in_array($arrPath[1], static::$fileExts) ) {
            throw new Exception("Unrecognized File. Exit process!", 400);
        }

        $this->reader = IOFactory::createReader(ucfirst($arrPath[1]));
        $this->reader->setReadDataOnly(TRUE);
        $this->spreadsheet = $this->reader->load($filePath);
    }

    public function validateAndResult()
    {
        $worksheet = $this->spreadsheet->getActiveSheet();
        // Get the highest row and column numbers referenced in the worksheet
        $highestRow = $worksheet->getHighestRow(); // e.g. 10
        $highestColumn = $worksheet->getHighestColumn(); // e.g 'F'
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);

        // row header index, we assume row header is 0
        $rowHeader = 1;

        // collect column header name and also check symbol
        $headers = [];
        for($col = 1; $col <= $highestColumnIndex; ++$col) {
            if( $worksheet->getCellByColumnAndRow($col, $rowHeader) ) {
                $d = $worksheet->getCellByColumnAndRow($col, $rowHeader)->getValue();
                if( $d ) {
                    if( str_contains($d, '*') ) {
                        $headers[$col] = ['CheckEmpty', $d];
                    } else if( str_contains($d, '#') ) {
                        $headers[$col] = ['CheckSpace', $d];
                    }
                }
            }
        }

        // Here we go collect errors every single row field
        $this->tables = [];
        for ($row = 2; $row <= (int)$highestRow; ++$row) {
            foreach($headers as $index => $mark) {
                if( $worksheet->getCellByColumnAndRow($index, $row, false) ) {
                    $a = $worksheet->getCellByColumnAndRow($index, $row, false)->getValue();
                    $a = utf8_decode($a);
                    if( ($mark[0] == 'CheckEmpty' && empty($a) ) ||
                        ($mark[0] == 'CheckSpace' && str_contains($a, " ") )
                    ) {
                       $this->fillArray($this->tables, $row, $mark);
                    }
                } else {
                    $this->fillArray($this->tables, $row, $mark);
                }
            }
        }

        // Return this class and next you decide what output are JSON, Array, table command
        return $this;
    }

    protected function fillArray(&$tables, $row, $mark)
    {
        if( isset($tables[$row]) ) {
            array_push( $tables[$row]['error'], $this->translate($mark[0], $mark[1]) );
        } else {
            $tables[$row] = [
                'row'   => $row,
                'error' => [$this->translate($mark[0], $mark[1])]
            ];
        }
    }

    protected function translate($key, $field)
    {
        if( $key == 'CheckEmpty' ) {
            $field = trim($field,"*");
            return "Missing value in ${field}";
        }

        if( $key == 'CheckSpace' ) {
            $field = trim($field,"#");
            return "${field} should not contain any space";
        }

        return "";
    }

    /**
     * Array to table command output
     */
    public function toTable()
    {
        if( empty($this->tables) ) {
            return [];
        }

        $_rows = [];
        foreach ($this->tables as $key => $value) {
            $_rows[] = [$value['row'], join(", ", $value['error']) ];
        }

        return $_rows;
    }
}
