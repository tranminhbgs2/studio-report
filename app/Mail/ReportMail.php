<?php

namespace App\Mail;

use DateTime;
use DateTimeZone;
use Illuminate\Bus\Queueable;
use Illuminate\Mail\Mailable;
use Illuminate\Queue\SerializesModels;

class ReportMail extends Mailable
{
    use Queueable, SerializesModels;

    public $filePaths;

    public function __construct($filePaths)
    {
        $this->filePaths = $filePaths;
    }

    public function build()
    {
        $dateTime = new DateTime();
        $dateTime->setTimezone(new DateTimeZone('Asia/Bangkok'));
        $dateTime->modify('-1 month');
        $currentMonth = $dateTime->format('m/Y');
        $email = $this->subject('Báo cáo lương Tổng hợp Tháng ' . $currentMonth)
            ->view('emails.report');

        foreach ($this->filePaths as $file) {
            // echo $file;die;
            $email->attachFromStorageDisk('public', $file);
        }

        return $email;
    }
}
