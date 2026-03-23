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

class EstudiosXServicioRXExport implements FromView, WithEvents, WithTitle
{
    public $mes;
    public $year;

    function __construct($mes, $year) { 
        $this->mes = $mes;
        $this->year = $year;
    }

    public function view(): View {
        return view('admin.appointments.reports.prueba', ['data' => []]);
    }

    public function title(): string {   
        return 'Estudios x Servicio';
    }

    public function registerEvents(): array {
        return [
            AfterSheet::class => function(AfterSheet $event) {
                $sheet = $event->sheet->getDelegate();
                
                // OPTIMIZACIÓN DE MEMORIA
                \PhpOffice\PhpSpreadsheet\Calculation\Calculation::getInstance($event->sheet->getParent())->disableCalculationCache();

                // 1. FORMATO Y ENCABEZADOS
                $sheet->getStyle('A1:AG112')->getBorders()->getAllBorders()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $sheet->getStyle('A1:AG112')->getAlignment()->setHorizontal('center')->setVertical('center');
                $sheet->getColumnDimension('A')->setWidth(250, 'px');
                $sheet->freezePane('B3');

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

                // 2. FUNCIÓN DE PROCESAMIENTO OPTIMIZADA
                // Esta función reemplaza los cientos de IFs por un mapeo directo en memoria
                $procesarBloque = function($parentId, $startRow, $label = null, $isSingleService = false) use ($sheet, $columnas) {
                    if($label) {
                        $sheet->setCellValue('A'.($startRow-1), $label);
                        $sheet->getStyle('A'.($startRow-1))->getFont()->setBold(true);
                    }

                    // Consulta maestra para el bloque actual
                    $query = DB::table('details_appointments')
                        ->select(
                            DB::raw('Day(appointments.date) AS dia'), 
                            'details_appointments.idservice', 
                            DB::raw('COUNT(details_appointments.id) AS total')
                        )
                        ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                        ->join('services', 'services.id', '=', 'details_appointments.idservice')
                        ->whereMonth('appointments.date', $this->mes)
                        ->whereYear('appointments.date', $this->year)
                        ->where('appointments.status', 3);

                    if($isSingleService) {
                        $query->where('services.id', $parentId);
                        $servicios = Service::where('id', $parentId)->where('status', 1)->get();
                    } else {
                        $query->where('services.parent_id', $parentId);
                        $servicios = Service::where('parent_id', $parentId)->where('status', 1)->get();
                        if($parentId == 4) {
                            $query->where('services.id', '<>', 63);
                            $servicios->where('id', '<>', 63);
                        }
                        $servicios = $servicios->get();
                    }

                    $datosMap = $query->groupBy('dia', 'details_appointments.idservice')->get()->groupBy('idservice');

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
                    return $currentRow;
                };

                // 3. EJECUCIÓN POR SECCIONES
                // Hospitalización (ID 1) -> Fila 3
                $procesarBloque(1, 3);
                $sheet->setCellValue('A20', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'20', "=SUM({$c}3:{$c}19)");

                // Consulta Externa (ID 2) -> Fila 21
                $procesarBloque(2, 21);
                $sheet->setCellValue('A55', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'55', "=SUM({$c}21:{$c}54)");

                // Emergencias (ID 3) -> Fila 57
                $procesarBloque(3, 57, 'EMERGENCIAS');
                $sheet->setCellValue('A64', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'64', "=SUM({$c}57:{$c}63)");

                // Unidades Externas (ID 63) -> Fila 66
                $procesarBloque(63, 66, 'SERVICIOS A OTRAS UNIDADES', true);
                $sheet->setCellValue('A67', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'67', "=SUM({$c}66:{$c}66)");

                // Apoyo (ID 4) -> Fila 70
                $procesarBloque(4, 70, 'SERVICIOS DE APOYO');
                $sheet->setCellValue('A111', 'SUB-TOTAL');
                foreach($columnas as $c) $sheet->setCellValue($c.'111', "=SUM({$c}70:{$c}110)");

                // 4. GRAN TOTAL (Fila 112)
                $sheet->setCellValue('A112', 'TOTAL');
                $sheet->getStyle('A112:AG112')->getFont()->setBold(true);
                foreach($columnas as $c) {
                    $sheet->setCellValue($c.'112', "={$c}20+{$c}55+{$c}64+{$c}67+{$c}111");
                }
                $sheet->setCellValue('AG112', "=SUM(B112:AF112)");
            },
        ];
    }
}