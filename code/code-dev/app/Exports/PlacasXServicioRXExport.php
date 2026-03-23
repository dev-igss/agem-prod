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

class PlacasXServicioExport implements FromView, WithEvents, WithTitle
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
        return 'Reporte de Placas RX';
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

                // Encabezados Estilo "Placas / Rayos X"
                $sheet->getColumnDimension('A')->setWidth(280, 'px');
                $sheet->mergeCells('B1:AF1');
                $sheet->setCellValue('A1', 'CONTROL DE PLACAS Y ESTUDIOS RADIOLÓGICOS');
                $sheet->setCellValue('A2', 'Estudio / Día del Mes');
                $sheet->getStyle('A1:AG2')->getFont()->setBold(true);
                $sheet->getStyle('A1:AG2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

                foreach ($columnas_datos as $index => $col) {
                    $sheet->getColumnDimension($col)->setWidth(35, 'px');
                    $sheet->setCellValue($col . '2', $index + 1);
                }
                $sheet->setCellValue($colTotal . '2', 'TOTAL');

                // 2. Definición de Secciones
                // Nota: Ajusta el parent_id según donde guardes los servicios de Rayos X
                $secciones = [
                    ['titulo' => 'RADIOLOGÍA E IMAGEN', 'parent_id' => 4], 
                ];

                foreach ($secciones as $sec) {
                    $sheet->setCellValue('A' . $this->currentRow, $sec['titulo']);
                    $sheet->getStyle('A' . $this->currentRow . ':AG' . $this->currentRow)->getFill()
                        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        ->getStartColor()->setRGB('CFE2F3'); // Azul claro para Placas
                    $this->currentRow++;

                    $startSectionRow = $this->currentRow;

                    // FILTRO: Servicios activos (status = 1)
                    $servicios = Service::where('status', 1)
                        ->where('parent_id', $sec['parent_id'])
                        // Si tienes una subcategoría específica para placas, puedes agregar otro filtro aquí
                        ->get();

                    $datos = DB::table('details_appointments')
                        ->select(
                            DB::raw('Day(appointments.date) AS dia'),
                            'services.id AS idservicio',
                            DB::raw('COUNT(details_appointments.id) AS total_placas') 
                        )
                        ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                        ->join('services', 'services.id', '=', 'details_appointments.idservice')
                        ->whereMonth('appointments.date', $this->mes)
                        ->whereYear('appointments.date', $this->year)
                        ->where('appointments.status', 3)
                        ->where('services.status', 1)
                        ->whereIn('services.id', $servicios->pluck('id'))
                        ->groupBy('dia', 'idservicio')
                        ->get()
                        ->groupBy('idservicio');

                    foreach ($servicios as $srv) {
                        $sheet->setCellValue('A' . $this->currentRow, $srv->name);
                        
                        if ($datos->has($srv->id)) {
                            foreach ($datos[$srv->id] as $registro) {
                                $col = $columnas_datos[$registro->dia - 1];
                                $sheet->setCellValue($col . $this->currentRow, $registro->total_placas);
                            }
                        }
                        $sheet->setCellValue($colTotal . $this->currentRow, "=SUM(B{$this->currentRow}:AF{$this->currentRow})");
                        $this->currentRow++;
                    }

                    // Fila de Total de Placas
                    $sheet->setCellValue('A' . $this->currentRow, 'TOTAL GENERAL DE PLACAS');
                    $sheet->getStyle('A' . $this->currentRow . ':AG' . $this->currentRow)->getFont()->setBold(true);
                    foreach (array_merge($columnas_datos, [$colTotal]) as $col) {
                        $sheet->setCellValue($col . $this->currentRow, "=SUM({$col}{$startSectionRow}:{$col}" . ($this->currentRow - 1) . ")");
                    }
                    $this->currentRow += 2;
                }

                // 3. Estética final
                $rangoFinal = "A1:AG" . ($this->currentRow - 2);
                $sheet->getStyle($rangoFinal)->applyFromArray([
                    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN, 'color' => ['rgb' => '000000']]],
                    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
                ]);

                $sheet->setShowGridlines(false);
                $sheet->freezePane('B3');
                $sheet->getProtection()->setSheet(true);
            },
        ];
    }
}