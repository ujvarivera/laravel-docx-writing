<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

class DocWriterController extends Controller
{
    public function index() 
    {
        try {
            $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor('sample.docx');
            $templateProcessor->setValue('ugyfelnev', 'Kiss Béla Zrt.');
            $templateProcessor->setValue('cim', '9400 Sopron, Új utca 52');
            $templateProcessor->setValue('vezeto', 'Kiss Béla');
            $templateProcessor->setValue('telefonszam', '+36325646644');
            $templateProcessor->saveAs('word.docx');

            return('Writing the doc file was successful');

        } catch (\Error $e) {
            abort(500, $e->getMessage());
        }
    }
}
