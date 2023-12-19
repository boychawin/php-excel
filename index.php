<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Html;
// use PhpOffice\PhpSpreadsheet\Style\Fill;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$spreadsheet->getProperties()->setCreator('Maarten Balliauw')
    ->setLastModifiedBy('Maarten Balliauw')
    ->setTitle('Office 2007 XLSX Test Document')
    ->setSubject('Office 2007 XLSX Test Document')
    ->setDescription('Test document for Office 2007 XLSX, generated using PHP classes.')
    ->setKeywords('office 2007 openxml php')
    ->setCategory('Test result file');


// Generate demo data for 10 rows
$data = [];

for ($i = 1; $i <= 10; $i++) {
    $innerData = []; // Initialize an empty array for each $i

    for ($j = 1; $j <= 110; $j++) {
        $innerData[] = $j; // Add $j to the inner array
    }

    $data[] = $innerData; // Add the inner array to $data
}



// Main header
$mainHeader = [
    'Transaction details', 'Ordering customer', 'Ordering customer contact details', 'Ordering customer business details',
    'Ordering customer identification details', 'Beneficiary customer', 'Beneficiary customer contact details',
    'Beneficiary customer business details', 'Beneficiary customer account details', 'Person/organisation accepting the transfer instruction from the ordering customer',
    'Person/organisation accepting the money or property from the ordering customer (if different)', 'Person/organisation sending the transfer instruction (if different)',
    'Person/organisation receiving the transfer instruction', 'Person/organisation distributing money or property (if different)',
    'Retail outlet/business location where money or property is being distributed (if different)', 'Reason', 'Person completing this report'
];

$subHeaders = [
    'Transaction details' => [
        'Date money/property received from the ordering customer', 'Date money/property made available to the beneficiary customer', 'Currency code',
        'Total amount/value', 'Type of transfer', 'Description of property', 'Transaction reference number'
    ],
    'Ordering customer' => [
        'Full name', 'If known by any other name', 'Date of birth (if an individual)'
    ],
    'Ordering customer contact details' => [
        'Business/residential address (not a post box address)', 'City/town/suburb', 'State', 'Postcode', 'Country',
        'Postal address', 'City/town/suburb', 'State', 'Postcode', 'Country', 'Phone', 'Email'
    ],
    'Ordering customer business details' => [
        'Occupation, business or principal activity', 'ABN, ACN or ARBN', 'Customer number (allocated by remitter)',
        'Account number (held by remitter)',
        'Business structure (if not an individual)'
    ],
    'Ordering customer identification details' => [
        'ID type (1)',
        'ID type (if \'Other\')',
        'Number', 'Issuer',
        'ID type (2)',
        'ID type (if \'Other\')',
        'Number',
        'Issuer',
        'Electronic data source'
    ],
    'Beneficiary customer' => [
        'Full name', 'Date of birth (if an individual)', 'Any business name under which the beneficiary customer is operating'
    ],
    'Beneficiary customer contact details' => [
        'Business/residential address (not a post box address)', 'City/town/suburb', 'State', 'Postcode', 'Country',
        'Postal address', 'City/town/suburb', 'State', 'Postcode', 'Country', 'Phone', 'Email'
    ],
    'Beneficiary customer business details' => [
        'Occupation, business or principal activity', 'ABN, ACN or ARBN', 'Business structure (if not an individual)'
    ],
    'Beneficiary customer account details' => [
        'Account number', 'Name of institution (where account is held)', 'City', 'Country'
    ],
    'Person/organisation accepting the transfer instruction from the ordering customer' => [
        'Identification number of the retail outlet/business location', 'Full name', 'Business/residential address (not a post box address)',
        'City/town/suburb', 'State', 'Postcode', 'Is this person/organisation accepting the money or property?',
        'Is this person/organisation sending the transfer instruction?'
    ],
    'Person/organisation accepting the money or property from the ordering customer (if different)' => [
        'Full name', 'Business/residential address (not a post box address)', 'City/town/suburb',
        'State', 'Postcode'
    ],
    'Person/organisation sending the transfer instruction (if different)' => [
        'Full name', 'If known by any other name', 'Date of birth (if an individual)', 'Business/residential address (not a post box address)',
        'City/town/suburb', 'State', 'Postcode', 'Postal address', 'City/town/suburb', 'State', 'Postcode',
        'Phone', 'Email', 'Occupation, business or principal activity', 'ABN, ACN or ARBN', 'Business structure (if not an individual)'
    ],
    'Person/organisation receiving the transfer instruction' => [
        'Full name', 'Business/residential address (not a post box address)', 'City/town/suburb', 'State',
        'Postcode', 'Country'
    ],
    'Person/organisation distributing money or property (if different)' => [
        'Full name', 'Business/residential address (not a post box address)', 'City/town/suburb', 'State',
        'Postcode', 'Country'
    ],
    'Retail outlet/business location where money or property is being distributed (if different)' => [
        'Full name', 'Business/residential address (not a post box address)', 'City/town/suburb', 'State',
        'Postcode', 'Country'
    ],
    'Reason' => [
        'Reason for the transfer'
    ],
    'Person completing this report' => [
        'Full name', 'Job title', 'Phone', 'Email'
    ]
];

// Set main headers in row 1



$columnStart = 'A';
$rowStart = 1;

foreach ($mainHeader as $header) {
    if (isset($subHeaders[$header])) {
        $subHeaderCount = count($subHeaders[$header]);
        $columnEnd = $columnStart;
        for ($i = 1; $i <= $subHeaderCount - 1; $i++) {
            $columnEnd++;
        }

        if (
            $header == 'Ordering customer' ||
            $header == 'Ordering customer business details' ||
            $header == 'Ordering customer identification details' ||
            $header == 'Person/organisation accepting the transfer instruction from the ordering customer' ||
            $header == 'Reason' ||
            $header == 'Person/organisation sending the transfer instruction (if different)' ||
            $header ==   "Person/organisation distributing money or property (if different)"
        ) {
            $color = "c9c299";
        } else {
            $color = "FFFFDD";
        }

        $endColumn = $columnEnd . $rowStart;
        $sheet->mergeCells($columnStart . $rowStart . ':' . $endColumn);
        $sheet->setCellValue($columnStart . $rowStart, $header);
        $sheet->getStyle($columnStart . $rowStart . ':' . $endColumn)
            ->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB($color);

        foreach ($subHeaders[$header] as $subHeader) {
            $sheet->setCellValue($columnStart . ($rowStart + 1), $subHeader);
            $sheet->getStyle($columnStart . ($rowStart + 1))->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB($color);
            $columnStart++;
        }
    } else {
        $sheet->setCellValue($columnStart . $rowStart, $header);
        $sheet->getStyle($columnStart . $rowStart)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('000000');
        $columnStart++;
    }
}




// // Set width for main headers
// $sheet->getColumnDimension('A')->setWidth(40); // Adjust width for the main header cell
// $column = 'B';
// foreach ($mainHeader as $header) {
//     $sheet->getColumnDimension($column)->setWidth(25); // Adjust width for main header cells
//     $column++;
// }


// Set width for Transaction details header columns (A to G) and align text left
// $columnsForTransactionDetails = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];
// foreach ($columnsForTransactionDetails as $col) {
//     $sheet->getColumnDimension($col)->setWidth(25); // Adjust width for Transaction details header columns
//     $sheet->getStyle($col . '2')->getAlignment()->setHorizontal('left'); // Align text left for Transaction details header
// }

// // Add demo data
$row = 3; // Start from the third row
foreach ($data as $rowData) {
    $column = 'A';
    foreach ($rowData as $value) {
        $sheet->setCellValue($column . $row, $value);
        $column++;
    }
    $row++;
}

$dataFooter[] = [
    'Created Date', 'Created Date', 'AUD', 'Send Amount', 'Money', 'Money', 'Money', 'Transaction Number', 'Sender Firstname + Lastname', '_',
    'Sender Date of Birth', 'Sender Address', 'Sender City', 'Sender State', 'Sender Postcode', 'Sender Country', '_', '_', '_', '_', '_',
    'Sender phone', 'Sender email', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_',
    'Receiver Firstname + Lastname', '_', '_', 'Receiver Address', 'Receiver City', 'Receiver State', 'Receiver postcode', 'Receiver Country', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_',
    'Reciver Account Number', 'Receiver Bank name', 'Receiver Bank City', 'Receiver Bank Country', '4', 'Acare Business Solutions Pty Ltd', 'Shop T26, Level 1, Capitol Square 730-742 George Street', 'Haymarket', 'NSW', '2000',
    'Yes', 'Yes', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_', '_',
    'Transferrer Name', 'Transferrer Address', 'Transferrer City', 'Transferrer State', 'Transferrer Postcode', 'Transferrer Country', 'Yes', 'No', '_', '_', '_', '_', '_', '_', '_', '_', '_',
    'Purpose of transfer', '_', '_', '_', '_'
];

$rowFooterStart = count($data) + 4;


foreach ($dataFooter as $rowData) {
    $column = 'A';

    $subHeaderCount = count($rowData);
    $columnEnd = $column; // Change this variable name from $columnStart to $column

    for ($i = 1; $i <= $subHeaderCount - 1; $i++) {
        $columnEnd++;
    }

    $endColumn = $columnEnd . $rowFooterStart;

    $sheet->getStyle($column . $rowFooterStart . ':' . $endColumn)
        ->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('ffff00');

    foreach ($rowData as $value) {
        $sheet->setCellValue($column . $rowFooterStart, $value);
        $column++;
    }

    $rowFooterStart++;
}





// Save the Excel file
$writer = new Xlsx($spreadsheet);
$filename = 'transaction_details.xlsx';
$writer->save($filename);

// Generate HTML preview
$htmlWriter = new Html($spreadsheet);
$htmlFilename = 'transaction_details.html';
$htmlWriter->save($htmlFilename);

// HTML content to display a message after creating the Excel file with a full-width preview
echo '<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Excel File Created</title>
    <style>
        iframe {
            width: 100%;
            height: 80vh;
            border: none;
        }
    </style>
</head>
<body>
    <h1>Excel file created successfully!</h1>
    <p>You can download the Excel file <a href="' . $filename . '">here</a>.</p>
    <h2>Preview:</h2>
    <iframe src="' . $htmlFilename . '"></iframe>
</body>
</html>';
