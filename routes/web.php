<?php
use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelMergerController;

Route::get('/', function () {
    return view('welcome');
});

Route::get('/merge-excel', [ExcelMergerController::class, 'mergeExcel']);
Route::get('/reformat-excel', [ExcelMergerController::class, 'reformatExcel']);
