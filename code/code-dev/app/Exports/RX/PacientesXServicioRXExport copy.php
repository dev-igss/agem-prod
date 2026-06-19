<?php

namespace App\Exports;

use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Concerns\WithDrawings;
use Maatwebsite\Excel\Concerns\WithTitle;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Maatwebsite\Excel\Concerns\WithMultipleSheets;

use App\Models\Service;
use Illuminate\Support\Facades\Http;
use DB, Carbon\Carbon,  Auth;

class PacientesXServicioRXExport implements FromView, WithEvents, WithTitle
{
    /**
    * @return \Illuminate\Support\Collection
    */
    public $mes;
    public $year;

    function __construct($mes, $year){ 
        $this->mes = $mes;
        $this->year = $year;
    }

    public function view(): View{
        

        $data = [
            
        ];

        return view('admin.appointments.reports.prueba', $data);
    }

    public function title(): string{   
        //return 'Día '.Carbon::now()->format('d');
        return 'Pacientes x Servicio';
    }

    public function registerEvents(): array{
        return [
            AfterSheet::class    => function(AfterSheet $event){
                /* Configuracion de la hoja */
                $event->getSheet()->getDelegate()->getStyle('A1:AG112')->applyFromArray(
                    array(
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                                'color' => ['rgb' => '000000'],
                            ],
                        ]
                    )
                );

                $event->sheet->getPageSetup()->setFitToPage(true);
                $event->sheet->setShowGridlines(False);

                $event->sheet->freezePane('B3');

                $event->sheet->getParent()->getActiveSheet()->getProtection()->setSheet(true);

                // lock all cells then unlock the cell
                $event->sheet->getParent()->getActiveSheet()
                    ->getStyle('B1:AF1')
                    ->getProtection()
                    ->setLocked(Protection::PROTECTION_UNPROTECTED);

                $event->sheet->getDelegate()->mergeCells('B1:AF1');

                /*Impresion de datos */
                $event->sheet->getDelegate()->getStyle('A1:AG125')
                                ->getAlignment()
                                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                
                $event->sheet->getDelegate()->getStyle('A1:AG125')
                                ->getAlignment()
                                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

                $event->sheet->getDelegate()->getColumnDimension('A')->setWidth(269, 'px');

                $event->sheet->setCellValue('A1', 'PACIENTES POR SERVICIO');              
                $event->sheet->getDelegate()->getStyle('A1')->getFont()->setBold(true);

                $event->sheet->setCellValue('A2', 'Días');              
                $event->sheet->getDelegate()->getStyle('A1')->getFont()->setBold(true);

                $dias_del_mes = [
                    '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
                    '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
                    '21', '22', '23', '24', '25', '26', '27', '28', '29', '30',
                    '31'
                ];

                $columnas_datos = [
                    'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 
                    'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
                    'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE',
                    'AF'
                ];

                /* Dias del mes */

                for($i = 0; $i < count($columnas_datos); $i++){
                    $event->sheet->getDelegate()->getColumnDimension($columnas_datos[$i])->setWidth(37, 'px');
                    $event->sheet->setCellValue($columnas_datos[$i].'2', $dias_del_mes[$i]);
                }

                /* Servicios de Hospitalizacion */
                $servicios_hosp = Service::where('parent_id', 1)->get();
                $conteo_pacientes_hosp = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('services.id AS idservicio'), 
                        DB::raw('services.name AS servicio'),  
                        DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('services', 'services.id', '=', 'details_appointments.idservice')
                    ->whereMonth('appointments.date', $this->mes)
                    ->whereYear('appointments.date', $this->year)
                    ->where('appointments.area', 0) 
                    ->where('appointments.status', 3)
                    ->where('services.parent_id', 1)
                    ->groupBy(DB::raw('Day(appointments.date)'), DB::raw('services.id'))
                    ->get();

                $row_count1 = 3;                
                $row_hosp_datos = 3;
                for($i=0; $i < count($servicios_hosp); $i++) {
                    $event->sheet->setCellValue('A'.$row_count1, $servicios_hosp[$i]->name);                                  
                    $row_count1++;
                }
                for($j =0; $j < count($servicios_hosp); $j++){

                    $contador_hosp = 0;
                    for($d = 0; $d < count($conteo_pacientes_hosp); $d++){
                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 1){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('B'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 2){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('C'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 3){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('D'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 4){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('E'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 5){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('F'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 6){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('G'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 7){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('H'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 8){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('I'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 9){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('J'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 10){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('K'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 11){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('L'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 12){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('M'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 13){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('N'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 14){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('O'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 15){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('P'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 16){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('Q'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 17){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('R'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 18){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('S'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 19){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('T'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 20){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('U'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 21){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('V'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 22){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('W'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 23){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('X'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 24){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('Y'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 25){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('Z'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 26){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('AA'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 27){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('AB'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 28){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('AC'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 29){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('AD'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 30){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('AE'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_hosp[$contador_hosp]->dia == 31){                               
                            if($conteo_pacientes_hosp[$contador_hosp]->idservicio == $servicios_hosp[$j]->id){
                                $event->sheet->setCellValue('AF'.$row_hosp_datos, $conteo_pacientes_hosp[$contador_hosp]->total_pacientes);
                            }                                
                        }

                        $contador_hosp++;
                    }

                    $row_hosp_datos++;
                    if($j == 16){
                        $row_hosp_datos = 3;             
                    }  
                }
                $event->sheet->setCellValue('A'.$row_count1, 'SUB-TOTAL');
                for($i = 0; $i < count($columnas_datos); $i++){
                    $event->sheet->setCellValue($columnas_datos[$i].'20', '=SUM('.$columnas_datos[$i].'3:'.$columnas_datos[$i].'19)');
                }

                /* Servicios de Consulta Externa */
                $servicios_coex = Service::where('parent_id', 2)->get();
                $conteo_pacientes_coex = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('services.id AS idservicio'), 
                        DB::raw('services.name AS servicio'), 
                        DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('services', 'services.id', '=', 'details_appointments.idservice')
                    ->whereMonth('appointments.date', $this->mes)
                    ->whereYear('appointments.date', $this->year)
                    ->where('appointments.area', 0)
                    ->where('appointments.status', 3)
                    ->where('services.parent_id', 2)
                    ->groupBy(DB::raw('Day(appointments.date)'), DB::raw('services.id'))
                    ->get();

                $row_count2 = $row_count1+1;
                $row_coex_datos = 21;                  
                for($i=0; $i < count($servicios_coex); $i++) {
                    $event->sheet->setCellValue('A'.$row_count2, $servicios_coex[$i]->name);                                  
                    $row_count2++;
                }
                for($j =0; $j < count($servicios_coex); $j++){

                    $contador_coex = 0;
                    for($d = 0; $d < count($conteo_pacientes_coex); $d++){
                        if($conteo_pacientes_coex[$contador_coex]->dia == 1){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('B'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 2){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('C'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 3){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('D'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 4){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('E'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 5){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('F'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 6){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('G'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 7){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('H'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 8){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('I'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 9){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('J'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 10){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('K'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 11){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('L'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 12){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('M'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 13){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('N'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 14){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('O'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 15){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('P'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 16){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('Q'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 17){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('R'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 18){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('S'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 19){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('T'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 20){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('U'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 21){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('V'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 22){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('W'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 23){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('X'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 24){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('Y'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 25){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('Z'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 26){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('AA'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 27){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('AB'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 28){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('AC'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 29){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('AD'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 30){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('AE'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_coex[$contador_coex]->dia == 31){                               
                            if($conteo_pacientes_coex[$contador_coex]->idservicio == $servicios_coex[$j]->id){
                                $event->sheet->setCellValue('AF'.$row_coex_datos, $conteo_pacientes_coex[$contador_coex]->total_pacientes);
                            }                                
                        }

                        $contador_coex++;
                    }

                    $row_coex_datos++;
                    if($j == 33){
                        $row_coex_datos = 21;             
                    }  
                }
                $event->sheet->setCellValue('A'.$row_count2, 'SUB-TOTAL');
                for($i = 0; $i < count($columnas_datos); $i++){
                    $event->sheet->setCellValue($columnas_datos[$i].'55', '=SUM('.$columnas_datos[$i].'21:'.$columnas_datos[$i].'54)');
                }

                /* Servicios de Emergencias */
                $row_separador1 = $row_count2+1;
                $event->sheet->getDelegate()->mergeCells('A'.$row_separador1.':AG'.$row_separador1);
                $event->sheet->setCellValue('A'.$row_separador1 , 'EMERGENCIAS');
                $servicios_emer = Service::where('parent_id', 3)->get();
                $conteo_pacientes_emer = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('services.id AS idservicio'), 
                        DB::raw('services.name AS servicio'), 
                        DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('services', 'services.id', '=', 'details_appointments.idservice')
                    ->whereMonth('appointments.date', $this->mes)
                    ->whereYear('appointments.date', $this->year)
                    ->where('appointments.area', 0)
                    ->where('appointments.status', 3)
                    ->where('services.parent_id', 3)
                    ->groupBy(DB::raw('Day(appointments.date)'), DB::raw('services.id'))
                    ->get();

                $row_count3 = $row_count2+2;          
                $row_emer_datos = 57;    
                for($i=0; $i < count($servicios_emer); $i++) {
                    $event->sheet->setCellValue('A'.$row_count3, $servicios_emer[$i]->name);                                  
                    $row_count3++;
                }
                for($j =0; $j < count($servicios_emer); $j++){

                    $contador_emer = 0;
                    for($d = 0; $d < count($conteo_pacientes_emer); $d++){
                        if($conteo_pacientes_emer[$contador_emer]->dia == 1){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('B'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 2){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('C'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 3){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('D'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 4){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('E'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 5){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('F'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 6){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('G'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 7){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('H'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 8){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('I'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 9){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('J'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 10){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('K'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 11){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('L'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 12){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('M'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 13){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('N'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 14){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('O'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 15){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('P'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 16){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('Q'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 17){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('R'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 18){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('S'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 19){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('T'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 20){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('U'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 21){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('V'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 22){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('W'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 23){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('X'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 24){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('Y'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 25){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('Z'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 26){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('AA'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 27){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('AB'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 28){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('AC'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 29){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('AD'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 30){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('AE'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_emer[$contador_emer]->dia == 31){                               
                            if($conteo_pacientes_emer[$contador_emer]->idservicio == $servicios_emer[$j]->id){
                                $event->sheet->setCellValue('AF'.$row_emer_datos, $conteo_pacientes_emer[$contador_emer]->total_pacientes);
                            }                                
                        }

                        $contador_emer++;
                    }

                    $row_emer_datos++;
                    if($j == 10){
                        $row_emer_datos = 57;             
                    }  
                }
                $event->sheet->setCellValue('A'.$row_count3, 'SUB-TOTAL');
                for($i = 0; $i < count($columnas_datos); $i++){
                    $event->sheet->setCellValue($columnas_datos[$i].'64', '=SUM('.$columnas_datos[$i].'57:'.$columnas_datos[$i].'63)');
                }

                /* Servicios de Unidades Externas */
                $row_separador2 = $row_count3+1;
                $event->sheet->getDelegate()->mergeCells('A'.$row_separador2.':AG'.$row_separador2);
                $event->sheet->setCellValue('A'.$row_separador2 , 'SERVICIOS A OTRAS UNIDADES');
                $servicios_unidades_externas = Service::where('id', 63)->get();
                $conteo_pacientes_unidades_externas = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('services.id AS idservicio'), 
                        DB::raw('services.name AS servicio'), 
                        DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('services', 'services.id', '=', 'details_appointments.idservice')
                    ->whereMonth('appointments.date', $this->mes)
                    ->whereYear('appointments.date', $this->year)
                    ->where('appointments.area', 0)
                    ->where('appointments.status', 3)
                    ->where('services.id', 63)
                    ->groupBy(DB::raw('Day(appointments.date)'), DB::raw('services.id'))
                    ->get();

                    $row_count4 = $row_count3+2;                
                    $row_unidades_externas_datos = 66;              

                for($i=0; $i < count($servicios_unidades_externas); $i++) {
                    $event->sheet->setCellValue('A'.$row_count4, $servicios_unidades_externas[$i]->name);                                  
                    $row_count4++;
                }
                for($j =0; $j < count($servicios_unidades_externas); $j++){

                    $contador_unidades_externas = 0;
                    for($d = 0; $d < count($conteo_pacientes_unidades_externas); $d++){
                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 1){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('B'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 2){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('C'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 3){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('D'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 4){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('E'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 5){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('F'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 6){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('G'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 7){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('H'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 8){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('I'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 9){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('J'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 10){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('K'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 11){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('L'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 12){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('M'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 13){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('N'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 14){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('O'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 15){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('P'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 16){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('Q'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 17){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('R'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 18){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('S'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 19){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('T'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 20){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('U'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 21){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('V'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 22){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('W'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 23){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('X'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 24){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('Y'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 25){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('Z'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 26){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('AA'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 27){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('AB'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 28){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('AC'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 29){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('AD'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 30){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('AE'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->dia == 31){                               
                            if($conteo_pacientes_unidades_externas[$contador_unidades_externas]->idservicio == $servicios_unidades_externas[$j]->id){
                                $event->sheet->setCellValue('AF'.$row_unidades_externas_datos, $conteo_pacientes_unidades_externas[$contador_unidades_externas]->total_pacientes);
                            }                                
                        }

                        $contador_unidades_externas++;
                    }

                    $row_unidades_externas_datos++;
                    if($j == 10){
                        $row_unidades_externas_datos = 57;             
                    }  
                }
                $event->sheet->setCellValue('A'.$row_count4, 'SUB-TOTAL');
                for($i = 0; $i < count($columnas_datos); $i++){
                    $event->sheet->setCellValue($columnas_datos[$i].'67', '=SUM('.$columnas_datos[$i].'66:'.$columnas_datos[$i].'66)');
                }

                /* Servicios de Apoyo */
                $row_separador3 = $row_count4+1;
                $event->sheet->getDelegate()->mergeCells('A'.$row_separador3.':AG'.$row_separador3);
                $event->sheet->setCellValue('A'.$row_separador3 , 'SERVICIOS DE APOYO');
                $row_separador4 = $row_count4+2;
                $event->sheet->getDelegate()->mergeCells('A'.$row_separador4.':AG'.$row_separador4);
                $event->sheet->setCellValue('A'.$row_separador4 , 'SERVICIOS DE LA UNIDAD');
                $servicios_apoyo = Service::where('parent_id', 4)->where('id', '<>', 63)->get();
                $conteo_pacientes_servicios_apoyo = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('services.id AS idservicio'), 
                        DB::raw('services.name AS servicio'), 
                        DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('services', 'services.id', '=', 'details_appointments.idservice')
                    ->whereMonth('appointments.date', $this->mes)
                    ->whereYear('appointments.date', $this->year)
                    ->where('appointments.area', 0)
                    ->where('appointments.status', 3)
                    ->where('services.parent_id', 4)
                    ->where('services.id', '<>', 63)
                    ->groupBy(DB::raw('Day(appointments.date)'), DB::raw('services.id'))
                    ->get();

                $row_count5 = $row_count4+3;   
                $row_servicio_apoyo_datos = 70;              
                for($i=0; $i < count($servicios_apoyo); $i++) {
                    $event->sheet->setCellValue('A'.$row_count5, $servicios_apoyo[$i]->name);                                  
                    $row_count5++;
                }
                for($j =0; $j < count($servicios_apoyo); $j++){

                    $contador_servicio_apoyo = 0;
                    for($d = 0; $d < count($conteo_pacientes_servicios_apoyo); $d++){
                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 1){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('B'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 2){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('C'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 3){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('D'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 4){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('E'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 5){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('F'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 6){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('G'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 7){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('H'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 8){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('I'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 9){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('J'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 10){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('K'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 11){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('L'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 12){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('M'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 13){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('N'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 14){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('O'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 15){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('P'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 16){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('Q'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 17){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('R'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 18){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('S'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 19){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('T'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 20){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('U'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 21){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('V'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 22){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('W'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 23){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('X'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 24){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('Y'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 25){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('Z'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 26){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('AA'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 27){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('AB'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 28){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('AC'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 29){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('AD'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 30){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('AE'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->dia == 31){                               
                            if($conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->idservicio == $servicios_apoyo[$j]->id){
                                $event->sheet->setCellValue('AF'.$row_servicio_apoyo_datos, $conteo_pacientes_servicios_apoyo[$contador_servicio_apoyo]->total_pacientes);
                            }                                
                        }

                        $contador_servicio_apoyo++;
                    }

                    $row_servicio_apoyo_datos++;
                    if($j == 40){
                        $row_servicio_apoyo_datos = 74;             
                    }   
                }
                $event->sheet->setCellValue('A'.$row_count5, 'SUB-TOTAL');
                for($i = 0; $i < count($columnas_datos); $i++){
                    $event->sheet->setCellValue($columnas_datos[$i].'111', '=SUM('.$columnas_datos[$i].'70:'.$columnas_datos[$i].'110)');
                }

                /*Gran Total */
                $row_count6 = $row_count5+1;
                $event->sheet->setCellValue('A'.$row_count6  , 'TOTAL');
                for($i = 0; $i < count($columnas_datos); $i++){
                    $event->sheet->setCellValue($columnas_datos[$i].$row_count6 , '='.$columnas_datos[$i].'20+'.$columnas_datos[$i].'55+'.$columnas_datos[$i].'64+'.$columnas_datos[$i].'67+'.$columnas_datos[$i].'111');
                }

                /*Total por servicios y subtotales*/
                $event->sheet->getDelegate()->mergeCells('AG1:AG2');
                $event->sheet->setCellValue('AG1' , 'TOTAL');
                for($i=3; $i<113; $i++){
                    $event->sheet->setCellValue('AG'.$i, '=SUM(B'.$i.':AF'.$i.')');
                }
            },
        ];
    }
}
