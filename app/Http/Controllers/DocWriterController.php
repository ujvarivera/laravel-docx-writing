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

    public function makeTable() {

        /*
        $sectionStyle = array(
            'orientation' => 'landscape',
            'marginTop' => 600,
            'colsNum' => 2,
        );
        */
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        //$section = $phpWord->addSection($sectionStyle);

        $section = $phpWord->addSection();
        $table = $section->addTable(array('borderSize' => 1, 'borderColor' => '999999', 'afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0 ));
        
        $styleCell = array('borderTopSize'=>1 ,'borderTopColor' =>'black','borderLeftSize'=>1,'borderLeftColor' =>'black','borderRightSize'=>1,'borderRightColor'=>'black','borderBottomSize' =>1,'borderBottomColor'=>'black' );
        $TfontStyle = array('bold'=>true, 'italic'=> true, 'size'=>11, 'name' => 'Times New Roman', 'afterSpacing' => 0, 'Spacing'=> 0, 'cellMargin'=>0);
        
        $table->addRow(-0.5, array('exactHeight' => -5));
        $cell = $table->addCell(2500, $styleCell)->addText('Hello', $TfontStyle, array('align' => 'left', 'spaceAfter' => 0));
        $cell2 = $table->addCell(2500, $styleCell)->addText('Bye', $TfontStyle, array('align' => 'left', 'spaceAfter' => 0));
        $cell3 = $table->addCell(2500, $styleCell)->addText('Hiii', $TfontStyle, array('align' => 'left', 'spaceAfter' => 0));

        $table->addRow();
        $table->addCell()->addText('Hello', $TfontStyle);
        $table->addCell()->addText('Bye', $TfontStyle);
        $table->addCell()->addText('Hiii');

        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save('helloWorld.docx');
        

    }
}
