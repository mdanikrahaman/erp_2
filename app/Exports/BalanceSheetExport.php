<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromArray;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\BeforeWriting;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Concerns\WithCustomStartCell;
use Illuminate\Support\Collection;

class BalanceSheetExport implements FromArray, WithEvents, WithHeadings, WithStyles, WithColumnWidths, WithCustomStartCell

{
    /**
    * @return \Illuminate\Support\Collection
    */

    public function __construct($data , $startDate, $endDate, $companyName)
    {
        $formattedData = [];
        foreach ($data  as $category => $subCategories)
        {
            $amountTotal= 0; $total =0;

            foreach ($subCategories as $subkey => $subCategory)
            {

            foreach($subCategory['account']  as $a)
            {
                if($a['totalDebit'] != null || $a['totalCredit'] != null)
                {
                    if($category == 'Liabilities' || $category == 'Equity')
                    {
                        $formattedData[] = [
                            'Account Name' => 'Liabilities & Equity',
                            'Account No'   => '',
                            'Total'        => ''
                        ];
                    }

                    if($category == 'Liabilities' || $category == 'Equity')
                    {
                        $formattedData[] = [
                            'Account Name' => '  ' .$category,
                            'Account No'   => '',
                            'Total'        => ''
                        ];
                    }
                    else
                    {
                        $formattedData[] = [
                            'Account Name' =>  $category,
                            'Account No'   => '',
                            'Total'        => ''
                        ];
                    }
                    break;
                }
            }
        }
            foreach ($subCategories as $subkey => $subCategory)
            {
                foreach($subCategory['account']  as $a)
                {
                    if($a['totalDebit'] != null || $a['totalCredit'] != null)
                    {
                        $formattedData[] = [
                            'Account Name' => '    ' . $subCategory['subType'],
                            'Account No'   => '',
                            'Total'        => ''
                        ];
                        break;
                    }
                }

                $debitTotal = 0;  

                foreach($subCategory['account']  as $a)
                {
                    if($a['totalDebit'] != null || $a['totalCredit'] != null)
                    {
                       
                        $formattedData[] = [
                            'Account Name' => '       ' .  $a['account_name'],
                            'Account No'   => $a['account_no'],
                            'Total'        => $a['netAmount']
                        ];

                        $acc = $subCategory['subType'];
                        $amountTotal += $a['netAmount'];
                    }
                }
                
                foreach($subCategory['account']  as $a)
                {
                    if($a['totalDebit'] != null || $a['totalCredit'] != null)
                    {
                        $formattedData[] = [
                            'Account Name' => '    Total ' . $subCategory['subType'],
                            'Account No'   => '',
                            'Total'        => $amountTotal
                        ];
                        break;
                    }
                }
            }
            if(($category == 'Liabilities' || $category == 'Equity') && $amountTotal != 0)
            {
                $formattedData[] = [
                    'Account Name' => '  Total ' .$category,
                    'Account No'   => '',
                    'Total'        => $amountTotal
                ];
            }
            else
            {
                if($amountTotal != 0)
                {
                    $formattedData[] = [
                        'Account Name' =>  'Total ' . $category,
                        'Account No'   => '',
                        'Total'        => $amountTotal
                    ];
                }
            }

            if(($category == 'Liabilities' || $category == 'Equity') && $amountTotal != 0)
            {
                foreach($formattedData as $a)
                {
                    ($a['Account Name'] == 'Total Liabilities' || $a['Account Name'] == 'Total Equity');
                }
                $total += $a['Total'];
                $formattedData[] = [
                    'Account Name' => 'Total Liabilities & Equity',
                    'Account No'   => '',
                    'Total'        => $total
                ];
                break;
            }
        }
        
        $formattedData;

        $this->data        = $formattedData;
        $this->companyName = $companyName;
        $this->startDate   = $startDate;
        $this->endDate     = $endDate;
    }

    public function startCell(): string
    {
        return 'A5';
    }

    public function columnWidths(): array
    {
        return [
            'A' => 30,
            'B' => 15,
            'C' => 15,
            'D' => 15,

        ];
    }

    public function styles(Worksheet $sheet)
    {
        $sheet->getStyle('A5')->getFont()->setBold(true);
        $sheet->getStyle('B5')->getFont()->setBold(true);
        $sheet->getStyle('C5')->getFont()->setBold(true);
        $sheet->getStyle('D5')->getFont()->setBold(true);
        $sheet->getStyle('E5')->getFont()->setBold(true);
        $sheet->getStyle('F5')->getFont()->setBold(true);
    }

    public function array(): array
    {
        return $this->data ;
    }



    public function registerEvents(): array
    {
        return [
            BeforeWriting::class => function (BeforeWriting $event) {

            },

            AfterSheet::class => function (AfterSheet $event) {
                $event->sheet->getDelegate()->mergeCells('A1:F1');
                $event->sheet->getDelegate()->mergeCells('A2:F2');
                $event->sheet->getDelegate()->mergeCells('A3:F3');

                $event->sheet->getDelegate()->setCellValue('A1', 'Balance Sheet - ' . $this->companyName)->getStyle('A1')->getFont()->setBold(true);
                $event->sheet->getDelegate()->setCellValue('A2', 'Print Out Date : ' . date('Y-m-d H:i'));
                $event->sheet->getDelegate()->setCellValue('A3', 'Date : ' . $this->startDate . ' - ' . $this->endDate);

                $event->sheet->getStyle('A1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $event->sheet->getStyle('A2')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $event->sheet->getStyle('A3')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

                $event->sheet->getDelegate()->getStyle('A')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $event->sheet->getDelegate()->getStyle('B')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

                $data = $this->data;
                foreach ($data as $index => $row) {

                    if (isset($row['Account Name']) && ($row['Account Name'] == '  Total Liabilities' || $row['Account Name'] == 'Assets' || 
                    $row['Account Name'] ==  'Total Assets' || $row['Account Name'] ==  '  Liabilities' || $row['Account Name'] ==  '  Equity' || 
                    $row['Account Name'] ==  '  Total Equity' || $row['Account Name'] == 'Liabilities & Equity' || $row['Account Name'] == 'Total Liabilities & Equity')) {
                        $rowIndex = $index + 6; // Adjust for 1-based indexing and header row
                        $event->sheet->getStyle('A' . $rowIndex . ':C' . $rowIndex)
                            ->applyFromArray([
                                'font' => [
                                    'bold' => true,
                                ],
                            ]);
                    }
                }

            },
        ];
    }

    public function headings(): array
    {
        return [
            "Account",
            "Account No",
            "Total",
        ];
    }
}
