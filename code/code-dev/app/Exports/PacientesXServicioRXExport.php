<?php

namespace App\Exports;

use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use App\Models\Service;
use Illuminate\Support\Facades\DB;
use Carbon\Carbon;

class PacientesXServicioRXExport implements FromView, WithEvents, WithTitle
{
    public $mes;
    public $year;
    private $currentRow = 3; 

    public function __construct($mes, $year)
    {
        $this->mes = $mes;
        $this->year = $year;
    }

    public function view(): View
    {
        return view('admin.appointments.reports.prueba', ['data' => []]);
    }

    public function title(): string
    {
        return 'Pacientes x Servicio';
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function(AfterSheet $event) {
                $sheet = $event->sheet->getDelegate();
                
                // 1. Configuración de Columnas
                $columnas_datos = [];
                for ($i = 0; $i < 31; $i++) {
                    $columnas_datos[] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i + 2);
                }
                $colTotal = 'AG';

                // Estilos de Encabezado
                $sheet->getColumnDimension('A')->setWidth(250, 'px');
                $sheet->mergeCells('B1:AF1');
                $sheet->setCellValue('A1', 'PACIENTES POR SERVICIO');
                $sheet->setCellValue('A2', 'Días');
                $sheet->getStyle('A1:AG2')->getFont()->setBold(true);
                $sheet->getStyle('A1:AG2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

                foreach ($columnas_datos as $index => $col) {
                    $sheet->getColumnDimension($col)->setWidth(35, 'px');
                    $sheet->setCellValue($col . '2', $index + 1);
                }
                $sheet->setCellValue($colTotal . '2', 'TOTAL');

                // 2. Definición de Secciones Dinámicas
                $secciones = [
                    ['titulo' => 'HOSPITALIZACIÓN', 'parent_id' => 1],
                    ['titulo' => 'CONSULTA EXTERNA', 'parent_id' => 2],
                    ['titulo' => 'EMERGENCIAS', 'parent_id' => 3],
                    ['titulo' => 'SERVICIOS A OTRAS UNIDADES', 'id' => 63],
                    ['titulo' => 'SERVICIOS DE APOYO', 'parent_id' => 4, 'exclude_id' => 63],
                ];

                $filasSubtotales = [];

                foreach ($secciones as $sec) {
                    // Título de Sección
                    $sheet->setCellValue('A' . $this->currentRow, $sec['titulo']);
                    $sheet->getStyle('A' . $this->currentRow . ':AG' . $this->currentRow)->getFill()
                        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        ->getStartColor()->setRGB('E9E9E9');
                    $this->currentRow++;

                    $startSectionRow = $this->currentRow;

                    // --- CONSULTA DE SERVICIOS CON STATUS = 1 ---
                    $query = Service::where('status', 1); // <--- Condicional de status = 1
                    
                    if (isset($sec['id'])) $query->where('id', $sec['id']);
                    if (isset($sec['parent_id'])) $query->where('parent_id', $sec['parent_id']);
                    if (isset($sec['exclude_id'])) $query->where('id', '<>', $sec['exclude_id']);
                    $servicios = $query->get();

                    // --- CONSULTA DE DATOS ---
                    $datos = DB::table('details_appointments')
                        ->select(
                            DB::raw('Day(appointments.date) AS dia'),
                            'services.id AS idservicio',
                            DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes')
                        )
                        ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                        ->join('services', 'services.id', '=', 'details_appointments.idservice')
                        ->whereMonth('appointments.date', $this->mes)
                        ->whereYear('appointments.date', $this->year)
                        ->where('appointments.status', 3)
                        ->where('services.status', 1) // <--- Refuerzo de status = 1 en el join
                        ->whereIn('services.id', $servicios->pluck('id'))
                        ->groupBy('dia', 'idservicio')
                        ->get()
                        ->groupBy('idservicio');

                    foreach ($servicios as $srv) {
                        $sheet->setCellValue('A' . $this->currentRow, $srv->name);
                        
                        if ($datos->has($srv->id)) {
                            foreach ($datos[$srv->id] as $registro) {
                                $col = $columnas_datos[$registro->dia - 1];
                                $sheet->setCellValue($col . $this->currentRow, $registro->total_pacientes);
                            }
                        }
                        $sheet->setCellValue($colTotal . $this->currentRow, "=SUM(B{$this->currentRow}:AF{$this->currentRow})");
                        $this->currentRow++;
                    }

                    // Fila de Subtotal
                    $sheet->setCellValue('A' . $this->currentRow, 'SUB-TOTAL ' . $sec['titulo']);
                    $sheet->getStyle('A' . $this->currentRow . ':AG' . $this->currentRow)->getFont()->setBold(true);
                    
                    foreach (array_merge($columnas_datos, [$colTotal]) as $col) {
                        $sheet->setCellValue($col . $this->currentRow, "=SUM({$col}{$startSectionRow}:{$col}" . ($this->currentRow - 1) . ")");
                    }
                    
                    $filasSubtotales[] = $this->currentRow;
                    $this->currentRow += 2; 
                }

                // 3. Gran Total Final
                $sheet->setCellValue('A' . $this->currentRow, 'TOTAL GENERAL');
                $sheet->getStyle('A' . $this->currentRow . ':AG' . $this->currentRow)->getFont()->setBold(true);
                foreach (array_merge($columnas_datos, [$colTotal]) as $col) {
                    $sumFormula = "=" . implode('+', array_map(fn($f) => "{$col}{$f}", $filasSubtotales));
                    $sheet->setCellValue($col . $this->currentRow, $sumFormula);
                }

                // 4. Estilos Finales
                $rangoFinal = "A1:AG" . $this->currentRow;
                $sheet->getStyle($rangoFinal)->applyFromArray([
                    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '000000']]],
                    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
                ]);

                $sheet->setShowGridlines(false);
                $sheet->freezePane('B3');
                $sheet->getProtection()->setSheet(true);
                $sheet->getStyle('B1:AF1')->getProtection()->setLocked(Protection::PROTECTION_UNPROTECTED);
            },
        ];
    }
}