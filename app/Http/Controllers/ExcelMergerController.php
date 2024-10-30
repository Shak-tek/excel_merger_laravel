<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Illuminate\Support\Facades\File;

class ExcelMergerController extends Controller
{
    public function mergeExcel()
    {
        $directoryPath = public_path('excels'); // Folder path where all Excel files are located
        $excelFiles = File::files($directoryPath);

        // Sort files by month name precedence
        usort($excelFiles, function ($fileA, $fileB) {
            return $this->getMonthOrder($fileA->getFilenameWithoutExtension()) <=> $this->getMonthOrder($fileB->getFilenameWithoutExtension());
        });

        $combinedQueries = [];
        $combinedPages = [];

        // Initialize file names array to create headers later
        $fileNames = array_map(function ($file) {
            return $file->getFilenameWithoutExtension();
        }, $excelFiles);
        $headerQueries = ["Top Queries"];
        $headerPages = ["Top Pages"];

        // Initialize the unified header
        foreach ($fileNames as $fileName){
            $headerQueries = array_merge($headerQueries,[$fileName." Clicks", $fileName." Impressions", $fileName." Position"]);
            $headerPages = array_merge($headerPages,[$fileName." Clicks", $fileName." Impressions", $fileName." Position"]);
        }

        foreach ($excelFiles as $file) {
            // Load each Excel file
            $spreadsheet = IOFactory::load($file->getPathname());

            // Read 'Queries' and 'Pages' sheets
            $queriesSheet = $spreadsheet->getSheetByName('Queries');
            $pagesSheet = $spreadsheet->getSheetByName('Pages');
            // Combine data from Queries and Pages sheets
            $combinedQueries = $this->combineSheets("Top Queries",$combinedQueries, $queriesSheet, $file->getFilenameWithoutExtension(),$headerQueries);
            $combinedPages = $this->combineSheets("Top Pages",$combinedPages, $pagesSheet, $file->getFilenameWithoutExtension(),$headerPages);
        }

        // Create a new spreadsheet and add the combined data
        $newSpreadsheet = new Spreadsheet();

        // Add Queries sheet
        $queriesSheet = $newSpreadsheet->createSheet(0);
        $queriesSheet->setTitle('Queries');
        $this->writeDataToSheet($queriesSheet, $combinedQueries);

        // Add Pages sheet
        $pagesSheet = $newSpreadsheet->createSheet(1);
        $pagesSheet->setTitle('Pages');
        $this->writeDataToSheet($pagesSheet, $combinedPages);

        // Save the new Excel file
        $writer = new Xlsx($newSpreadsheet);
        $writer->save(public_path('CombinedData.xlsx'));

        return response()->download(public_path('CombinedData.xlsx'));
    }

    private function getMonthOrder($fileName)
    {
        // Assuming the filename includes the month name in a format like "January_2023.xlsx"
        // Extracting the month from the filename
        preg_match('/(January|February|March|April|May|June|July|August|September|October|November|December)/i', $fileName, $matches);
        if (isset($matches[0])) {
            $month = strtolower($matches[0]); // convert to lower case for uniformity
            $months = [
                'january' => 1,
                'february' => 2,
                'march' => 3,
                'april' => 4,
                'may' => 5,
                'june' => 6,
                'july' => 7,
                'august' => 8,
                'september' => 9,
                'october' => 10,
                'november' => 11,
                'december' => 12,
            ];
            return $months[$month] ?? 0; // Return 0 if month is not found
        }
        return 0; // Default if no month is found
    }

    private function combineSheets($sheetHeaderName, $existingData, $sheet, $fileName, $currentFileHeader)
    {
        // Initialize headers if it's the first file
        if (empty($existingData)) {
            $existingData[] = $currentFileHeader; // Add the header for the first file
        }

        $sheetData = $sheet->toArray();
        $dataMap = [];

        // Skip header row and process the rest
        foreach ($sheetData as $key => $row) {
            if ($key === 0) continue; // Skip header row

            $query = $row[0]; // First column (Queries)
            if (!isset($dataMap[$query])) {
                // Initialize an entry for the query
                $dataMap[$query] = array_fill(0, 4, null); // Fill with null for all metrics
            }

            // Set the respective values for the current file
            $dataMap[$query][0] = $row[1] ?? ''; // Clicks
            $dataMap[$query][1] = $row[2] ?? ''; // Impressions
            $dataMap[$query][2] = $row[3] ?? ''; // Position
        }

        // Add the unique queries and their metrics to the existing data
        foreach ($dataMap as $query => $metrics) {
            $row = [$query]; // Start with the unique query
            $row = array_merge($row, $metrics); // Add metrics from the current file

            // Check if the query already exists in existingData
            $existingQueryIndex = array_search($query, array_column($existingData, 0));

            if ($existingQueryIndex !== false) {
                // If the query exists, update its metrics
                $existingData[$existingQueryIndex] = array_merge($existingData[$existingQueryIndex], $metrics);
            } else {
                // If the query does not exist, add a new row
                $existingData[] = $row;
            }
        }

        // Ensure to include the headers for subsequent files
        if (!in_array($currentFileHeader, $existingData)) {
            $existingData[] = $currentFileHeader;
        }

        return $existingData;
    }

    private function writeDataToSheet($sheet, $data)
    {
        foreach ($data as $rowIdx => $row) {
            foreach ($row as $colIdx => $value) {
                $sheet->setCellValueByColumnAndRow($colIdx + 1, $rowIdx + 1, $value);
            }
        }
    }

    function reformatExcel()
    {
        $inputFilePath = public_path('CombinedData.xlsx');
        $outputFilePath = public_path('ReformattedCombinedData.xlsx');
        // Load the existing spreadsheet
        $spreadsheet = IOFactory::load($inputFilePath);

        // Create a new Spreadsheet object for output
        $newSpreadsheet = new Spreadsheet();

        // Iterate through each sheet in the original spreadsheet
        foreach ($spreadsheet->getAllSheets() as $sheetIndex => $sheet) {
            // Get the headers and data
            $headers = $sheet->toArray()[0]; // Fetch the first row as headers
            $data = $sheet->toArray(null, true, true, false);

            // Define the new column order
            $clickColumns = [];
            $impressionColumns = [];
            $positionColumns = [];

            // Separate headers based on column types
            foreach ($headers as $key => $header) {
                if (strpos($header, 'Clicks') !== false) {
                    $clickColumns[$key] = $header;
                } elseif (strpos($header, 'Impressions') !== false) {
                    $impressionColumns[$key] = $header;
                } elseif (strpos($header, 'Position') !== false) {
                    $positionColumns[$key] = $header;
                }
            }

            // Create the new header row
            $newHeaders = array_merge(['Top Queries'], array_values($clickColumns), array_values($impressionColumns), array_values($positionColumns));

            // Create a new sheet for the current sheet
            $newSheet = $newSpreadsheet->createSheet($sheetIndex);
            $newSheet->setTitle($sheet->getTitle());

            // Write the new headers
            $newSheet->fromArray($newHeaders, null, 'A1');

            // Write the data with rearranged columns
            for ($row = 1; $row <= count($data); $row++) {
                $newRowData = [];

                // Add "Top Queries" value
                $newRowData[] = $data[$row][0] ?? ''; // Assuming "Top Queries" is in the first column

                // Add Clicks, Impressions, and Position data in the new order
                foreach ($clickColumns as $col => $header) {
                    $newRowData[] = $data[$row][$col] ?? ''; // Access data for Clicks
                }
                foreach ($impressionColumns as $col => $header) {
                    $newRowData[] = $data[$row][$col] ?? ''; // Access data for Impressions
                }
                foreach ($positionColumns as $col => $header) {
                    $newRowData[] = $data[$row][$col] ?? ''; // Access data for Positions
                }

                // Write to the new sheet
                $newSheet->fromArray($newRowData, null, 'A' . ($row + 1)); // Adjusting row index for output
            }
        }

        // Save the reformatted spreadsheet
        $writer = new Xlsx($newSpreadsheet);
        $writer->save($outputFilePath);
    }




}
