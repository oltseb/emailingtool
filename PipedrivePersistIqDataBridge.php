<?php

require 'vendor/autoload.php';
require_once "ValidateEmail.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Writer;

class EmailSplitter
{

    public $iterator = 0;

    public $newSheetData = array();

    protected $reader;

    public function __construct()
    {
        $this->reader = new Reader();
    }

    /**
     * @param $rowData
     * @return array|null
     */
    function splitEmails($rowData)
    {
        $ind = 3;
        if (!isset($rowData[$ind])) {
            return null;
        }
        $emails = preg_split("/[\s,;\n]+/", $rowData[$ind]);
        $newData = array();
        foreach ($emails as $key => $email) {
            $emailValidator = new ValidateEmail();
            $email = trim($email, " \n;.");
            echo "\n" . " Validating " . $email . " ...";

            try {
                $emailResult = $emailValidator->validate(array($email), "leksa.ukr@gmail.com");
            } catch (Exception $e) {
                echo "\n" . $email . " failed to validate, was not saved";
                $this->iterator++;
                continue;
            }

            if (!$emailResult[$email]) {
                echo "\n" . $email . " does not exist";
                $this->iterator++;
                echo "\n" . $this->iterator . " was processed";
                continue;
            }
            echo "\n" . $email . " was validated";
            foreach ($rowData as $index => $cellData) {
                if ($index === $ind) {
                    $newData[$key][$index] = $email;
                    continue;
                }
                $newData[$key][$index] = $cellData;
            }
            $this->iterator++;
            echo "\n" . $this->iterator . " was processed";
        }
        return $newData;
    }

    /**
     * @param $data
     * @param $workSheetPostfix
     * @param $rowId
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    function writeIntoFile($data, $workSheetPostfix, $rowId)
    {
        try {
            $spreadSheet = $this->reader->load('email-splitting-' . $workSheetPostfix . '.xlsx');
        } catch (Exception $exception) {
            echo "\n" . $exception->getMessage();
            $spreadSheet = new Spreadsheet();
        }
        $sheet = $spreadSheet->getActiveSheet();
        $firstLetter = 'A';
        for ($i = 0; $i < 4; $i++) {
            $sheet->setCellValue($firstLetter . $rowId, $data[$i]);
            $firstLetter++;
        }
        echo "\n" . "Writing row #" . $rowId;
        $writer = new Writer($spreadSheet);
        // Name of the output file
        $writer->save('email-splitting' . '-' . $workSheetPostfix . '.xlsx');
    }
}

/** @var EmailSplitter $splitter */
$splitter = new EmailSplitter();
$reader = new Reader();

//
// ***
// *** ***
// ***
//
// Entry data file
$spreadSheet = $reader->load('email-splitting-test.xlsx');
//
// ***
// *** ***
// ***
//

ini_set('max_execution_time', 3000);
ini_set('memory_limit', '3000M');

foreach ($spreadSheet->getWorksheetIterator() as $worksheet) {
    $newSheetData = array();
    $rowId = 1;
    foreach ($worksheet->getRowIterator() as $row) {
        $rowData = array();
        foreach ($row->getCellIterator() as $cell) {
            if ($cell !== null) {
                $rowData[] = $cell->getValue();
            }
        }
        $splitEmails = $splitter->splitEmails($rowData);
        if (!$splitEmails) {
            continue;
        }
        foreach ($splitEmails as $newDataRow) {
            $splitter->writeIntoFile($newDataRow, "validated", $rowId);
            $rowId++;
        }
    }
}