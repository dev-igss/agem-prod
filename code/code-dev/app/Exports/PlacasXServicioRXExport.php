<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use Illuminate\Contracts\View\View;
use App\Models\Service;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use DB, Carbon\Carbon;

class PlacasXServicioRXExport implements FromView, WithEvents, WithTitle
{
    public $mes;
    public $year;

    function __construct($mes, $year) { 
        $this->mes = $mes;
        $this->year = $year;
    }

    public function view(): View {
        // Retornamos una vista vacía o mínima ya que todo el peso está en el AfterSheet
        return view('admin.appointments.reports.prueba', ['data' => []]);
    }

    public function title(): string {   
        return 'Placas x Servicio';
    }

    public function registerEvents(): array {
        return [
            AfterSheet::class => function(AfterSheet $event) {
                $sheet = $event->sheet->getDelegate();
                
                // OPTIMIZACIÓN DE MOTOR: Desactivar caché de cálculos para ahorrar RAM en Docker
                \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($event->sheet->getParent())->disableCalculationCache();

                // 1. CONFIGURACIÓN INICIAL
                $sheet->getStyle('A1:AG125')->getAlignment()->setHorizontal('center')->setVertical('center');
                $sheet->getColumnDimension('A')->setWidth(250, 'px');
                $sheet->freezePane('B3');

                // Encabezados
                $sheet->setCellValue('A1', 'PACIENTES POR SERVICIO');
                $sheet->mergeCells('B1:AF1');
                $sheet->setCellValue('A2', 'Días');
                
                $columnas = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF'];
                foreach($columnas as $index => $col) {
                    $sheet->getColumnDimension($col)->setWidth(35, 'px');
                    $sheet->setCellValue($col.'2', $index + 1);
                }
                $sheet->setCellValue('AG1', 'TOTAL');
                $sheet->mergeCells('AG1:AG2');

                // --- LÓGICA DE PROCESAMIENTO DINÁMICO ---
                
                // Función auxiliar para procesar bloques (Hosp, Coex, Emer, etc.)
                $procesarBloque = function($parentId, $startRow, $label = null) use ($sheet, $columnas) {
                    if($label) $sheet->setCellValue('A'.($startRow-1), $label);

                    // Consulta única por bloque
                    $query = DB::table('details_appointments')
                        ->select(
                            DB::raw('Day(appointments.date) AS dia'), 
                            'details_appointments.idservice', 
                            DB::raw('SUM(materials_appointments.amount) AS total')
                        )
                        ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                        ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'details_appointments.idappointment')
                        ->join('services', 'services.id', '=', 'details_appointments.idservice')
                        ->whereMonth('appointments.date', $this->mes)
                        ->whereYear('appointments.date', $this->year)
                        ->where('appointments.status', 3) // Solo finalizadas
                        ->where('services.parent_id', $parentId);
                    
                    if($parentId == 4) $query->where('services.id', '<>', 63); // Excluir unidad externa del bloque apoyo

                    $datosMap = $query->groupBy('dia', 'details_appointments.idservice')->get()->groupBy('idservice');
                    $servicios = Service::where('parent_id', $parentId);
                    if($parentId == 4) $servicios->where('id', '<>', 63);
                    $servicios = $servicios->get();

                    $currentRow = $startRow;
                    foreach($servicios as $srv) {
                        $sheet->setCellValue('A'.$currentRow, $srv->name);
                        if($datosMap->has($srv->id)) {
                            foreach($datosMap->get($srv->id) as $reg) {
                                $sheet->setCellValue($columnas[$reg->dia - 1].$currentRow, $reg->total);
                            }
                        }
                        $sheet->setCellValue('AG'.$currentRow, "=SUM(B{$currentRow}:AF{$currentRow})");
                        $currentRow++;
                    }
                    return $currentRow; // Retorna la siguiente fila libre
                };

                // 2. EJECUCIÓN DE BLOQUES
                $next = $procesarBloque(1, 3); // Hospitalización
                $sheet->setCellValue('A20', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'20', "=SUM({$c}3:{$c}19)");

                $next = $procesarBloque(2, 21); // Consulta Externa
                $sheet->setCellValue('A55', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'55', "=SUM({$c}21:{$c}54)");

                $sheet->setCellValue('A56', 'EMERGENCIAS');
                $next = $procesarBloque(3, 57); // Emergencias
                $sheet->setCellValue('A64', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'64', "=SUM({$c}57:{$c}63)");

                // Bloque Unidades Externas (ID 63 manual)
                $sheet->setCellValue('A65', 'SERVICIOS A OTRAS UNIDADES');
                $sheet->setCellValue('A66', 'UNIDAD EXTERNA');
                // (Consulta simplificada para este servicio único)
                $ext = DB::table('details_appointments')
                        ->select(DB::raw('Day(appointments.date) AS dia'), DB::raw('SUM(materials_appointments.amount) AS total'))
                        ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                        ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'details_appointments.idappointment')
                        ->whereMonth('appointments.date', $this->mes)->whereYear('appointments.date', $this->year)
                        ->where('details_appointments.idservice', 63)->where('appointments.status', 3)
                        ->groupBy('dia')->get();
                foreach($ext as $e) $sheet->setCellValue($columnas[$e->dia -1].'66', $e->total);
                $sheet->setCellValue('AG66', "=SUM(B66:AF66)");
                $sheet->setCellValue('A67', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'67', "=SUM({$c}66:{$c}66)");

                $sheet->setCellValue('A68', 'SERVICIOS DE APOYO');
                $next = $procesarBloque(4, 70); // Apoyo
                $sheet->setCellValue('A111', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'111', "=SUM({$c}70:{$c}110)");

                // GRAN TOTAL
                $sheet->setCellValue('A112', 'TOTAL');
                foreach($columnas as $c) {
                    $sheet->setCellValue($c.'112', "={$c}20+{$c}55+{$c}64+{$c}67+{$c}111");
                }
                $sheet->setCellValue('AG112', "=SUM(B112:AF112)");

                // 3. SECCIÓN TAMAÑO DE PLACAS (Consulta Maestra de Materiales)
                $sheet->setCellValue('A115', 'UTILIZADAS DEL TAMAÑO');
                $materialesData = DB::table('materials_appointments')
                    ->select('material', DB::raw('Day(appointments.date) as dia'), DB::raw('SUM(amount) as total'))
                    ->join('appointments', 'appointments.id', '=', 'materials_appointments.idappointment')
                    ->whereMonth('appointments.date', $this->mes)
                    ->whereYear('appointments.date', $this->year)
                    ->where('appointments.status', 3)
                    ->groupBy('material', 'dia')
                    ->get()
                    ->groupBy('material');

                $labelsPlacas = [0 => '8*10', 1 => '10*12', 2 => '11*14', 3 => '14*17'];
                $rowPlaca = 116;
                foreach($labelsPlacas as $idMat => $label) {
                    $sheet->setCellValue('A'.$rowPlaca, $label);
                    if($materialesData->has($idMat)) {
                        foreach($materialesData->get($idMat) as $reg) {
                            $sheet->setCellValue($columnas[$reg->dia - 1].$rowPlaca, $reg->total);
                        }
                    }
                    $sheet->setCellValue('AG'.$rowPlaca, "=SUM(B{$rowPlaca}:AF{$rowPlaca})");
                    $rowPlaca++;
                }

                // Bordes finales
                $sheet->getStyle('A1:AG112')->getBorders()->getAllBorders()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $sheet->getStyle('A115:AG120')->getBorders()->getAllBorders()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            },
        ];
    }
}