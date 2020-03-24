<?php

namespace App\Http\Controllers;

use App\Http\Requests\TaalSpellingDocumentRequest;
use Illuminate\Http\Request;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style\ListItem;

class DocxController extends Controller
{
    public function generateTaalSpellingDocument(Request $request)
    {
        $phpWord = new PhpWord;
        $section = $phpWord->addSection();
        $section->addTitle($request->get('title'));
        $numberOfQuestions = $request->get('numberOfQuestions');


        $phpWord->addTableStyle('answerList', [
            'borderColor' => '000000',
            'borderSize' => 6,
            'cellMargin' => 50
        ]);
        $table = $section->addTable('answerList');
        $table->addRow();
        $firstRowStyle = ['bold' => true];
        $table->addCell()->addText('Vraag', $firstRowStyle);
        $table->addCell(5000)->addText('Antwoord', $firstRowStyle);
        for ($i = 1; $i <= $numberOfQuestions; $i++) {
            $table->addRow();
            $table->addCell()->addText("$i.");
            $table->addCell();
        }

        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save(storage_path('Taalspelling.docx'));
        return response()->download(storage_path('Taalspelling.docx'))->deleteFileAfterSend();
    }


    public function generateRekenenDocument(Request $request)
    {
        $phpWord = new PhpWord;
        $section = $phpWord->addSection();
        $text = $section->addTitle($request->get('title'));
        $numberOfQuestions = $request->get('numberOfQuestions');
        $phpWord->addTableStyle('answerList', [
            'borderColor' => '000000',
            'borderSize' => 6,
            'cellMargin' => 50,
            'textSize' => 8,
        ]);
        $table = $section->addTable('answerList');
        $table->addRow();
        $firstRowStyle = ['bold' => true];
        $table->addCell()->addText('Vraag', $firstRowStyle);
        $table->addCell(5000)->addText('Antwoord', $firstRowStyle);
        for ($i = 1; $i <= $numberOfQuestions; $i++) {
            $table->addRow();
            $table->addCell()->addText("$i.");
            $table->addCell();
        }

        $section->addText('');
        $phpWord->addTableStyle('wordList', [
            'borderColor' => '000000',
            'borderSize' => 6,
            'cellMargin' => 50
        ], [
            'font' => ['bold' => true],
        ]);
        $table = $section->addTable('wordList');
        $table->addRow();
        $table->addCell()->addText('Woord', $firstRowStyle);
        $table->addCell(5000)->addText('Betekenis woord', $firstRowStyle);
        $table->addRow();
        $table->addCell()->addText($request->get('word'));
        $table->addCell()->addText($request->get('wordMeaning'));

        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save(storage_path('Rekenen.docx'));
        return response()->download(storage_path('Rekenen.docx'))->deleteFileAfterSend();
    }
}
