<?php

namespace App\Exports;

use DateTime;
use DateTimeZone;
use Illuminate\Support\Facades\App;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\WithColumnFormatting;
use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithMapping;
use NumberFormatter;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class SchedulesExport implements FromCollection, WithHeadings, WithStyles, WithColumnFormatting, ShouldAutoSize, WithEvents
{
    protected $fixedSalary;
    protected $employeeId;
    protected $type;
    protected $totalSalary;

    public function __construct($fixedSalary = 5000000, $employeeId = "", $type = "", $totalSalary = 0) // Giả sử lương cứng là 5,000,000đ
    {
        $this->fixedSalary = $fixedSalary;
        $this->employeeId = $employeeId;
        $this->type = $type;
        $this->totalSalary = $totalSalary;
    }
    public function collection()
    {
        $firestore = App::make('firebase.firestore');
        $collectionReference = $firestore->collection('works');

        // Giả sử bạn chỉ muốn xuất dữ liệu của tháng hiện tại
        $currentMonthStart = (new DateTime('first day of previous month'));
        $currentMonthEnd = (new DateTime('last day of previous month'));

        // Thêm điều kiện lọc dữ liệu theo tháng (ví dụ, 'shootingDate' nằm trong khoảng tháng hiện tại)
        // Lưu ý: Điều này cần được thay đổi tùy theo cấu trúc dữ liệu và yêu cầu cụ thể của bạn
        $field = $this->checkTypeEmployee($this->type);
        $documents = $collectionReference->where('shootingDate', '>=', $currentMonthStart)
            ->where('shootingDate', '<=', $currentMonthEnd)
            ->where($field, '=', $this->employeeId)
            ->documents();

        $data = collect([]);

        foreach ($documents as $document) {
            if ($document->exists()) {
                $docData = $document->data();
                $docData['shootingDate'] = $docData['shootingDate'] ?: '';
                $dateTime = new DateTime($docData['shootingDate'], new DateTimeZone('UTC'));
                $dateTime->setTimezone(new DateTimeZone('Asia/Bangkok'));

                $vietnameseDayOfWeek = $this->getVietnameseDayOfWeek($dateTime->format('l'));

                $locations = $docData['locations'] ?? [];
                $address = $this->formatAddress($locations);
                $fieldPrice = $this->priceEmployee($this->type);
                $typeEmployee = $this->checkEmployee($this->type);
                $salary = $docData[$fieldPrice] ?? 0;
                $this->totalSalary += $salary;
                $data->push([
                    'Ngày' => $dateTime->format('d'),
                    'Tháng' => $dateTime->format('m'),
                    'Thứ' => $vietnameseDayOfWeek,
                    'Khách hàng' => $docData['customerName'] ?? '',
                    'Giờ' => $docData['shootingHour'] . ':' . $docData['shootingMinute'],
                    'Địa điểm' => $address,
                    'Nhân viên' => $docData[$typeEmployee]['name'] ?? '',
                    'Lương' => $salary,
                ]);
            }
        }
        return $data;
    }

    public function headings(): array
    {
        return ['Ngày', 'Tháng', 'Thứ', 'Khách hàng', 'Giờ', 'Địa điểm', 'Nhân viên', 'Lương'];
    }
    public function styles(Worksheet $sheet)
    {
        return [
            // Căn giữa tiêu đề
            1    => ['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]],
            // Căn giữa các cột Ngày, Tháng, và Thứ
            'A' => ['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]],
            'B' => ['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]],
            'C' => ['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]],
        ];
    }

    public function columnFormats(): array
    {
        return [
            // Định dạng tiền tệ cho cột Lương
            'H' => '#,##0₫',
        ];
    }
    private function getVietnameseDayOfWeek($englishDayOfWeek)
    {
        $mapping = [
            'Monday' => 'Thứ Hai',
            'Tuesday' => 'Thứ Ba',
            'Wednesday' => 'Thứ Tư',
            'Thursday' => 'Thứ Năm',
            'Friday' => 'Thứ Sáu',
            'Saturday' => 'Thứ Bảy',
            'Sunday' => 'Chủ Nhật',
        ];

        return $mapping[$englishDayOfWeek] ?? false;
    }

    private function checkTypeEmployee($type)
    {
        $mapping = [
            'MAKEUP' => 'makeupArtistId',
            'PHOTO' => 'photographerId',
            'DESIGNER' => 'designerId',
            'LETAN' => 'letanId',
            'CSKH' => 'cskhId',
        ];

        return $mapping[$type];
    }

    private function priceEmployee($type)
    {
        $mapping = [
            'MAKEUP' => 'makeupPrice',
            'PHOTO' => 'photographerPrice',
            'DESIGNER' => 'designerPrice',
            // 'LETAN' => 'cskhPrice',
            'CSKH' => 'cskhPrice',
        ];

        return $mapping[$type];
    }
    private function checkEmployee($type)
    {
        $mapping = [
            'MAKEUP' => 'makeupArtist',
            'PHOTO' => 'photographer',
            'DESIGNER' => 'designer',
            // 'LETAN' => 'cskhPrice',
            'CSKH' => 'cskh',
        ];

        return $mapping[$type];
    }

    private function formatAddress($locations)
    {
        return collect($locations)->pluck('name')->implode(', ');
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $sheet = $event->sheet;
                $highestRow = $sheet->getHighestRow(); // Lấy hàng cuối cùng
                $highestColumn = $sheet->getHighestColumn();
                // Áp dụng border
                $cellRange = 'A1:' . $highestColumn . $highestRow; // Điều chỉnh vùng dữ liệu
                $sheet->getStyle($cellRange)->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => '191e21'],
                        ],
                    ],
                ]);

                // Áp dụng màu nền
                $sheet->getStyle($cellRange)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('bbe3cb');
                $sheet->getStyle($cellRange)->getFont()->setBold(true);
                // Thêm thông tin "Lương Cứng" vào hàng cuối cùng
                $sheet->setCellValue('G' . ($highestRow + 1), 'Lương Cứng');
                $sheet->setCellValue('H' . ($highestRow + 1), $this->fixedSalary);

                // Cài đặt định dạng cho "Lương Cứng"
                $sheet->getStyle('H' . ($highestRow + 1))->getNumberFormat()
                    ->setFormatCode('"#,##0₫"');

                // Căn giữa label "Lương Cứng"
                $sheet->getStyle('G' . ($highestRow + 1))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                // Cài đặt định dạng cho "Lương Cứng"
                $sheet->getStyle('H' . ($highestRow + 1))->getNumberFormat()
                    ->setFormatCode('#,##0₫');
                // Áp dụng màu nền
                $sheet->getStyle('G' . ($highestRow + 1) . ':' . 'H' . ($highestRow + 1))->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('e2b56e');
                $sheet->getStyle('G' . ($highestRow + 1) . ':' . 'H' . ($highestRow + 1))->getFont()->setBold(true);
                $sheet->getStyle('G' . ($highestRow + 1) . ':' . 'H' . ($highestRow + 1))->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => '191e21'],
                        ],
                    ],
                ]);

                // Tính toán và thêm thông tin "Tổng"
                $totalRow = $highestRow + 2; // Hàng cho "Tổng" sẽ nằm ngay dưới "Lương Cứng"
                $totalSalary = $this->totalSalary + $this->fixedSalary; // Công thức tính tổng lương
                $sheet->setCellValue('G' . $totalRow, 'Tổng');
                $sheet->setCellValue('H' . $totalRow, $totalSalary);

                // Định dạng cho "Lương Cứng" và "Tổng"
                $sheet->getStyle('H' . ($totalRow) . ':H' . $totalRow)->getNumberFormat()
                    ->setFormatCode('#,##0₫');
                $sheet->getStyle('G' . ($totalRow) . ':G' . $totalRow)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                // Áp dụng màu nền
                $sheet->getStyle('G' . ($totalRow) . ':' . 'H' . ($totalRow))->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('e2b56e');

                $sheet->getStyle('G' . ($totalRow) . ':' . 'H' . ($totalRow))->getFont()->setBold(true);
                $sheet->getStyle('G' . ($totalRow) . ':' . 'H' . ($totalRow))->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => '191e21'],
                        ],
                    ],
                ]);
            },
        ];
    }
}
