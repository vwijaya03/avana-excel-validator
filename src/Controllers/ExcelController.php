<?php
namespace App\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelController
{
    public function validate($filePath, $docType)
    {
        try {
            $reader = IOFactory::createReader($docType);
            $spreadsheet = $reader->load($filePath);
        
            $worksheet = $spreadsheet->getActiveSheet();
            // Get the highest row number and column letter referenced in the worksheet
            $highestRow = $worksheet->getHighestRow(); // e.g. 10
            $data = $worksheet->getCellByColumnAndRow(1, 4)->getValue(); // e.g. 10
            $highestColumn = $worksheet->getHighestColumn(); // e.g 'F'
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
        
            $header = [];
            $errorMessage = '';
            $errorRow = 0;
            $error = [];
            $isError = false;
        
            echo '<table border=1>' . "\n";
            for ($row = 1; $row <= $highestRow; ++$row) {
                echo '<tr>' . PHP_EOL;
                for ($col = 1; $col < $highestColumnIndex; ++$col) {
                    $value = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
                    
                    if ($row == 1) {
                        array_push($header, $value);
                    }
        
                    if (strpos($header[$col-1], '*') !== false) {
                        if ($value == '' || $value == null) {
                            // echo '<td>' . 'Error value kosong' . '</td>' . PHP_EOL;
                            echo '<td>' . $value . '</td>' . PHP_EOL;
        
                            if ($isError) {
                                $errorMessage .= 'missing value in ' . $header[$col - 1] . ', ';
                                $error[count($error) - 1] = $row.'-'.$errorMessage;
                            } else {
                                $isError = true;
                                $errorMessage .= 'missing value in ' . $header[$col - 1] . ', ';
                                array_push($error, $row.'-'.$errorMessage);
                            }
                        } else {
                            echo '<td>' . $value . '</td>' . PHP_EOL;
                        }
                    } 
                    else if (strpos($header[$col-1], '#') !== false) {
                        if (strpos($value, ' ') !== false) {
                            // echo '<td>' . 'Error ada spasi' . '</td>' . PHP_EOL;
                            echo '<td>' . $value . '</td>' . PHP_EOL;
        
                            if ($isError) {
                                $errorMessage .= $header[$col - 1]. ' should not contains space, ';
                                $error[count($error) - 1] = $row.'-'.$errorMessage;
                            } else {
                                $isError = true;
                                $errorMessage .= $header[$col - 1]. ' should not contains space, ';
                                array_push($error, $row.'-'.$errorMessage);
                            }
                        } else {
                            echo '<td>' . $value . '</td>' . PHP_EOL;
                        }
                    } 
                    else {
                        echo '<td>' . $value . '</td>' . PHP_EOL;
                    }
                }
        
                $isError = false;
                $errorMessage = '';
                
                echo '</tr>' . PHP_EOL;
            }
            echo '</table>' . PHP_EOL;
        
            echo '<br><br> <table border=1>' . "\n";
        
            echo '<th>Row</th>';
            echo '<th>Error</th>';
            for ($i = 0; $i < count($error); $i++) {
                echo '<tr>' . PHP_EOL;
                    echo '<td>'. explode("-", $error[$i])[0] . '</td>' . PHP_EOL;
                    echo '<td>' . explode("-", $error[$i])[1] . '</td>' . PHP_EOL;
                echo '</tr>' . PHP_EOL;        
            }
            echo '</table>' . PHP_EOL;
        } catch(Exception $e) {
            die($e);
        }
    }
}