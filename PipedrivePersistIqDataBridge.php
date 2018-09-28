<?php

require 'vendor/autoload.php';
require_once "ValidateEmail.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Writer;

class EmailSplitter
{

    public $iterator = 0;

    /**
     * @param $rowData
     * @return array|null
     */
    function splitEmails($rowData) {
        $ind = 3;
        if (!isset($rowData[$ind]))
        {
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
     * @param $workSheetIndex
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    function writeIntoFile($data, $workSheetIndex) {

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        foreach ($data as $rowLineNumber => $rowData) {
            $firstLetter = 'A';
            $rowLineNumber += 1;

            for ($i = 0; $i < 4; $i++) {
                $sheet->setCellValue($firstLetter . $rowLineNumber, $rowData[$i]);
                $firstLetter++;
            }
        }
        $writer = new Writer($spreadsheet);
        // Name of the output file
        $writer->save('email-splitting' . '-' . $workSheetIndex . '.xlsx');
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
ini_set('memory_limit','3000M');

foreach ($spreadSheet->getWorksheetIterator() as $workSheetIndex => $worksheet) {
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
            $newSheetData[] = $newDataRow;
        }
    }
    $splitter->writeIntoFile($newSheetData, $workSheetIndex+41);
    $rowId++;
}