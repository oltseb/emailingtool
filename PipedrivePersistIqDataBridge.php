<?php

require 'vendor/autoload.php';
require_once "ValidateEmail.php";
require_once "Helper.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Reader;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as Writer;

class EmailSplitter
{

//    /** SHOULD EXIST IN ANOTHER CLASS */
//     * @var
//     *
//     * Filename from where to read
//     */
//    protected $inputFileName;
//
//    /**
//     * @var
//     *
//     * Filename to where to write
//     */
//    protected $outputFileName;
//
//    /**
//     * @var
//     *
//     * number of column, which stores email address
//     */
//    protected $emailColumn;
//
//    const INPUT_FILE_ARGUMENT = 1;
//
//    const EMAIL_COL_ARGUMENT = 2;
//
//    const OUTPUT_FILE_ARGUMENT = 3;

    public $iterator = 0;

    public $newSheetData = array();

    /**
     * @var Reader
     */
    protected $reader;

    /**
     * @var Helper
     */
    public $helper;

    protected $emailColumnIndex = 1;

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
        $ind = $this->emailColumnIndex;
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
                $emailResult = array();
                $emailResult[$email] = true;
                $emailResult = $emailValidator->validate(array($email), "leksa.ukr@gmail.com");
            } catch (Exception $e) {
                echo "\n" . $email . " failed to validate, was not saved";
                $this->iterator++;
                continue;
            }

            try {
                if (!isset($emailResult[$email])) {
                    throw new Exception("Wrong ditch");
                }
                if (!$emailResult[$email]) {
                    echo "\n" . $email . " does not exist";
                    $this->iterator++;
                    echo "\n" . $this->iterator . " was processed";
                    continue;
                }
            } catch (Exception $e)
            {
                echo "\n" . $email . " broken pipe during procession. Still saving into the sheet";
                $this->iterator++;
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
        $outputFile = 'newsletter-part-10' . '-' . $workSheetPostfix . '.xlsx';
        try {
            $spreadSheet = $this->reader->load($outputFile);
        } catch (Exception $exception) {
            echo "\n" . $exception->getMessage();
            $spreadSheet = new Spreadsheet();
        }
        $sheet = $spreadSheet->getActiveSheet();
        $firstLetter = 'A';
        for ($i = 0; $i < 2; $i++) {
            $sheet->setCellValue($firstLetter . $rowId, $data[$i]);
            $firstLetter++;
        }
        echo "\n" . "Writing row #" . $rowId;
        $writer = new Writer($spreadSheet);
        // Name of the output file
        $writer->save($outputFile);
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
$spreadSheet = $reader->load('to-work-with/newsletter-part-10.xlsx');
//
// ***
// *** ***
// ***
//

ini_set('max_execution_time', 30000);
ini_set('memory_limit', '30000M');

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