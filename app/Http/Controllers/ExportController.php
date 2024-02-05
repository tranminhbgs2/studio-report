<?php

namespace App\Http\Controllers;

use App\Exports\SchedulesExport;
use App\Mail\ReportMail;
use DateTime;
use DateTimeZone;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\App;
use Illuminate\Support\Facades\Mail;
use Illuminate\Support\Facades\Storage;

class ExportController extends Controller
{
    public function exportSchedules(Request $request)
    {
        $firestore = App::make('firebase.firestore');
        $employeesCollection = $firestore->collection('employees');
        $employees = $employeesCollection->where('status', '=', 'ACTIVATED')->documents();
        $listOfFiles = []; // Khởi tạo mảng để lưu trữ đường dẫn file
        $dateTime = new DateTime();
        $dateTime->setTimezone(new DateTimeZone('Asia/Bangkok'));
        $dateTime->modify('-1 month');
        $currentMonth = $dateTime->format('Y-m'); // Định dạng 'Năm-Tháng', ví dụ: '2023-03'
        $directory = 'exports/' . $currentMonth; // Đường dẫn thư mục cần tạo

        // Kiểm tra nếu thư mục chưa tồn tại và tạo thư mục mới
        Storage::disk('public')->makeDirectory($directory, 0775, true, true);

        foreach ($employees as $employee) {
            if ($employee->exists()) {
                $employeeData = $employee->data();
                $fixedSalary = $employeeData['salary'] ?? 0;

                $export = new SchedulesExport($fixedSalary, $employee->id(), $employeeData['type']);
                $fileName = $directory . '/Lương_' . $employeeData['name'] . '-Tháng-' . $currentMonth . '.xlsx';
                Excel::store($export, $fileName, 'public');
                $listOfFiles[] = $fileName;
            }
        }

        // Gửi email với danh sách các file
        Mail::to('trantuyen3721@gmail.com')->send(new ReportMail($listOfFiles));

        // Trả về phản hồi tùy chỉnh hoặc tải xuống file cuối cùng (tùy vào yêu cầu của bạn)
        return response()->json(['message' => 'Exported successfully', 'files' => $listOfFiles]);
    }
}
