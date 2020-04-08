<?php

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/store-to-server', function () {

    Excel::create('tester', function($excel) {
        $excel->sheet('test-sheet', function($sheet) {
        });
    })->store('xls');
    
});

Route::get('/save-to-browser', function () {

    Excel::create('tester', function($excel) {
        $excel->sheet('test-sheet', function($sheet) {
        });
    })->export('xls');
    
});

Route::get('/load-from-template', function () {

    $file = storage_path('templates/templates.xlsx');

    Excel::load($file, function($excel) {

        // Sheet1 is your sheet name
        $excel->sheet('Sheet1', function($sheet) {

            for ($i=0; $i < 100000; $i++) { 

                // copy style from templates
                $sheet->duplicateStyle(
                    $sheet->getStyle(
                        'A2:C2'),'A' . ($i + 2) . ':C' . ($i + 2)
                    );

                $sheet->SetCellValue('A' . ($i + 2), 'test ' . $i);
                $sheet->SetCellValue('B' . ($i + 2), 'test ' . $i);
                $sheet->SetCellValue('C' . ($i + 2), 'test ' . $i);
                
            }
        });
    })->export('xls');
    
});
