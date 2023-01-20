<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

class DocWriterController extends Controller
{
    public function index() 
    {
        try {
            if (auth()->check()) {

                $customer = auth()->user()->customer;

                $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor('sample.docx');
                $templateProcessor->setValue('customername', $customer->name);
                $templateProcessor->setValue('address', $customer->address);
                $templateProcessor->setValue('leader', $customer->leader);
                $templateProcessor->setValue('phone', $customer->phone);
                $templateProcessor->saveAs('doc_' . $customer->id .'.docx');
    
                return('Writing the doc file was successful');
            }
            abort(403, 'Please login to continue!');

        } catch (\Error $e) {
            abort(500, $e->getMessage());
        }
    }
}
