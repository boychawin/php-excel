<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Style\Fill;

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
// $data = [];
// for ($i = 1; $i <= 10; $i++) {
//     $data[] = [
//         "Transaction $i", "Customer $i", "Contact $i", "Business $i", "Identification $i",
//         "Beneficiary $i", "Beneficiary Contact $i", "Beneficiary Business $i", "Account $i",
//         "Accepting Person/Org $i", "Accepting Money/Org $i", "Sending Person/Org $i",
//         "Receiving Person/Org $i", "Distributing Person/Org $i", "Retail Outlet $i",
//         "Reason $i", "Person Completing $i"
//     ];
// }

// Main header
$mainHeader = [
    'Transaction details', 'Ordering customer', 'Ordering customer contact details', 'Ordering customer business details',
    'Ordering customer identification details', 'Beneficiary customer', 'Beneficiary customer contact details',
    'Beneficiary customer business details', 'Beneficiary customer account details', 'Person/organisation accepting the transfer instruction from the ordering customer',
    'Person/organisation accepting the money or property from the ordering customer (if different)', 'Person/organisation sending the transfer instruction (if different)',
    'Person/organisation receiving the transfer instruction', 'Person/organisation distributing money or property (if different)',
    'Retail outlet/business location where money or property is being distributed (if different)', 'Reason', 'Person completing this report'
];

// Sub-headers for 'Transaction details'
$subHeadersTransactionDetails = [
    'Date money/property received from the ordering customer', 'Date money/property made available to the beneficiary customer', 'Currency code',
    'Total amount/value', 'Type of transfer', 'Description of property', 'Transaction reference number'
];

$subHeadersOrderingCustomer = [
    'Full name', 'If known by any other name', 'Date of birth (if an individual)'
];

// Sub-headers for 'Ordering customer contact details'
$subHeadersOrderingCustomerContact = [
    'Business/residential address (not a post box address)', 'City/town/suburb', 'State', 'Postcode', 'Country',
    'Postal address', 'City/town/suburb', 'State', 'Postcode', 'Country', 'Phone', 'Email'
];

// Sub-headers for 'Ordering customer business details'
$subHeadersOrderingCustomerBusiness = [
    'Occupation, business or principal activity', 'ABN, ACN or ARBN', 'Customer number (allocated by remitter)',
    'Account number (held by remitter)', 'Business structure (if not an individual)'
];

// Sub-headers for 'Ordering customer identification details'
$subHeadersOrderingCustomerIdentification = [
    'ID type (1)', 'ID type (if \'Other\')', 'Number', 'Issuer', 'ID type (2)', 'ID type (if \'Other\')', 'Number', 'Issuer', 'Electronic data source'
];

// Sub-headers for 'Beneficiary customer'
$subHeadersBeneficiaryCustomer = [
    'Full name', 'Date of birth (if an individual)', 'Any business name under which the beneficiary customer is operating'
];
$subHeadersBeneficiaryCustomerContact = [
    'Business/residential address (not a post box address)', 'City/town/suburb', 'State', 'Postcode', 'Country',
    'Postal address', 'City/town/suburb', 'State', 'Postcode', 'Country', 'Phone', 'Email'
];
$subHeadersBeneficiaryCustomerBusiness = [
    'Occupation, business or principal activity', 'ABN, ACN or ARBN', 'Business structure (if not an individual)'
];

$subHeadersBeneficiaryCustomerAccount = [
    'Account number', 'Name of institution (where account is held)', 'City', 'Country'
];
$subHeadersAcceptingTransfer = [
    'Identification number of the retail outlet/business location', 'Full name', 'Business/residential address (not a post box address)',
    'City/town/suburb', 'State', 'Postcode', 'Is this person/organisation accepting the money or property?',
    'Is this person/organisation sending the transfer instruction?'
];
$subHeadersAcceptingMoney = [
    'Full name', 'Business/residential address (not a post box address)', 'City/town/suburb',
    'State', 'Postcode'
];

$subHeadersSendingInstruction = [
    'Full name', 'If known by any other name', 'Date of birth (if an individual)', 'Business/residential address (not a post box address)',
    'City/town/suburb', 'State', 'Postcode', 'Postal address', 'City/town/suburb', 'State', 'Postcode',
    'Phone', 'Email', 'Occupation, business or principal activity', 'ABN, ACN or ARBN', 'Business structure (if not an individual)'
];

$subHeadersDistributingMoney = [
    'Full name', 'Business/residential address (not a post box address)', 'City/town/suburb', 'State',
    'Postcode', 'Country'
];

$subHeadersRetailOutlet = [
    'Full name', 'Business/residential address (not a post box address)', 'City/town/suburb', 'State',
    'Postcode', 'Country'
];

$subHeadersReason = [
    'Reason for the transfer'
];
$subHeadersCompletingReport = [
    'Full name', 'Job title', 'Phone', 'Email'
];

// Set main headers in row 1
$column = 'A';
$index = 1;

foreach ($mainHeader as $header) {
    if ($header == "Transaction details") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersTransactionDetails) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersTransactionDetails) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Transaction details' below the main header
        foreach ($subHeadersTransactionDetails as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } elseif ($header == "Ordering customer") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomer) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomer) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Ordering customer' below the main header
        foreach ($subHeadersOrderingCustomer as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } elseif ($header == "Ordering customer contact details") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomerContact) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomerContact) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Ordering customer contact details' below the main header
        foreach ($subHeadersOrderingCustomerContact as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }

        
    } elseif ($header == "Ordering customer business details") {
    //     $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomerBusiness) - 1) . '1');
    //     $sheet->setCellValue($column . '1', $header);
    //     $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomerBusiness) - 1) . '1')
    //         ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

    //     // Set sub-headers for 'Ordering customer business details' below the main header
    //     foreach ($subHeadersOrderingCustomerBusiness as $subHeader) {
    //         $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
    //         $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
    //         $column++;
    //     }
    } elseif ($header == "Ordering customer identification details") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomerIdentification) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersOrderingCustomerIdentification) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Ordering customer identification details' below the main header
        foreach ($subHeadersOrderingCustomerIdentification as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } elseif ($header == "Beneficiary customer") {
    //     $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomer) - 1) . '1');
    //     $sheet->setCellValue($column . '1', $header);
    //     $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomer) - 1) . '1')
    //         ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

    //     // Set sub-headers for 'Beneficiary customer' below the main header
    //     foreach ($subHeadersBeneficiaryCustomer as $subHeader) {
    //         $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
    //         $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
    //         $column++;
    //     }
    } elseif ($header == "Beneficiary customer contact details") {
    //     $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomerContact) - 1) . '1');
    //     $sheet->setCellValue($column . '1', $header);
    //     $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomerContact) - 1) . '1')
    //         ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

    //     // Set sub-headers for 'Beneficiary customer contact details' below the main header
    //     foreach ($subHeadersBeneficiaryCustomerContact as $subHeader) {
    //         $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
    //         $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
    //         $column++;
    //     }
    } elseif ($header == "Beneficiary customer business details") {
    //     $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomerBusiness) - 1) . '1');
    //     $sheet->setCellValue($column . '1', $header);
    //     $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomerBusiness) - 1) . '1')
    //         ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

    //     // Set sub-headers for 'Beneficiary customer business details' below the main header
    //     foreach ($subHeadersBeneficiaryCustomerBusiness as $subHeader) {
    //         $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
    //         $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
    //         $column++;
    //     }
    } elseif ($header == "Beneficiary customer account details") {
    //     $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomerAccount) - 1) . '1');
    //     $sheet->setCellValue($column . '1', $header);
    //     $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersBeneficiaryCustomerAccount) - 1) . '1')
    //         ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

    //     // Set sub-headers for 'Beneficiary customer account details' below the main header
    //     foreach ($subHeadersBeneficiaryCustomerAccount as $subHeader) {
    //         $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
    //         $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
    //         $column++;
    //     }
    } elseif ($header == "Person/organisation accepting the transfer instruction from the ordering customer") {
    //     $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersAcceptingTransfer) - 1) . '1');
    //     $sheet->setCellValue($column . '1', $header);
    //     $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersAcceptingTransfer) - 1) . '1')
    //         ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

    //     // Set sub-headers for 'Person/organisation accepting the transfer instruction' below the main header
    //     foreach ($subHeadersAcceptingTransfer as $subHeader) {
    //         $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
    //         $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
    //         $column++;
    //     }
    } elseif ($header == "Person/organisation accepting the money or property from the ordering customer (if different)") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersAcceptingMoney) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersAcceptingMoney) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Person/organisation accepting the money or property' below the main header
        foreach ($subHeadersAcceptingMoney as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } elseif ($header == "Person/organisation sending the transfer instruction (if different)") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersSendingInstruction) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersSendingInstruction) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Person/organisation sending the transfer instruction' below the main header
        foreach ($subHeadersSendingInstruction as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } elseif ($header == "Person/organisation distributing money or property (if different)") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersDistributingMoney) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersDistributingMoney) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Person/organisation distributing money or property' below the main header
        foreach ($subHeadersDistributingMoney as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } elseif ($header == "Retail outlet/business location where money or property is being distributed (if different)") {
    //     $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersRetailOutlet) - 1) . '1');
    //     $sheet->setCellValue($column . '1', $header);
    //     $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersRetailOutlet) - 1) . '1')
    //         ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

    //     // Set sub-headers for 'Retail outlet/business location' below the main header
    //     foreach ($subHeadersRetailOutlet as $subHeader) {
    //         $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
    //         $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
    //         $column++;
    //     }
    } elseif ($header == "Reason") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersReason) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersReason) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Reason' below the main header
        foreach ($subHeadersReason as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } elseif ($header == "Person completing this report") {
        $sheet->mergeCells($column . '1:' . chr(ord($column) + count($subHeadersCompletingReport) - 1) . '1');
        $sheet->setCellValue($column . '1', $header);
        $sheet->getStyle($column . '1:' . chr(ord($column) + count($subHeadersCompletingReport) - 1) . '1')
            ->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');

        // Set sub-headers for 'Person completing this report' below the main header
        foreach ($subHeadersCompletingReport as $subHeader) {
            $sheet->setCellValue($column . '2', $subHeader); // Set sub-headers in row 2
            $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
            $column++;
        }
    } else {
        $cell = $column . $index;
        $sheet->setCellValue($cell, $header);
        $sheet->getStyle($cell)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('000000');
        $column++;
    }
}

// // Set width for main headers
// $sheet->getColumnDimension('A')->setWidth(40); // Adjust width for the main header cell
// $column = 'B';
// foreach ($mainHeader as $header) {
//     $sheet->getColumnDimension($column)->setWidth(25); // Adjust width for main header cells
//     $column++;
// }

// Set subheaders in row 2
// $column = 'A';
// foreach ($subHeadersTransactiondetails as $subHeader) {
//     $sheet->setCellValue($column . '2', $subHeader);
//     // Set background color for subheader cell
//     $sheet->getStyle($column . '2')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFDD');
//     $sheet->getStyle($column . '2')->getAlignment()->setHorizontal('left'); // Align text left for subheader cell
//     $column++;
// }

// Set width for Transaction details header columns (A to G) and align text left
// $columnsForTransactionDetails = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];
// foreach ($columnsForTransactionDetails as $col) {
//     $sheet->getColumnDimension($col)->setWidth(25); // Adjust width for Transaction details header columns
//     $sheet->getStyle($col . '2')->getAlignment()->setHorizontal('left'); // Align text left for Transaction details header
// }

// // Add demo data
// $row = 3; // Start from the third row
// foreach ($data as $rowData) {
//     $column = 'A';
//     foreach ($rowData as $value) {
//         $sheet->setCellValue($column . $row, $value);
//         $column++;
//     }
//     $row++;
// }

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
