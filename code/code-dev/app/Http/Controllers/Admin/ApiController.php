<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Http;

use App\Models\SettingHolyDays, App\Models\Patient, App\Models\CodePatient, App\Models\Appointment, App\Models\ControlAppointment;
use App\Models\DetailAppointment, App\Models\Service, App\Models\Studie, App\Models\Schedule;
use Carbon\Carbon, DB;

use Artisaninweb\SoapWrapper\SoapWrapper;


class ApiController extends Controller
{
   
    public function getPatient($code, $exam){
        $patient = Patient::where('affiliation', $code)
            ->orderBy('type', 'ASC')
            ->get(); 

        if(count($patient) == 0):
            $data = [
                
            ];
        else:
            if(count($patient) > 1):
                foreach($patient as $p):

                    if(  $p->type == '0' || $p->type == '3' || $p->type == '4' || $p->type == '5' || $p->type == '6'):
                        $id_temp = $p->id;
                    endif;
                endforeach;
            else:
                foreach($patient as $p):
                    $id_temp = $p->id;
                endforeach;
            endif;

            switch($exam):
                case 0:
                    $code_temp_last = CodePatient::where('patient_id', $id_temp)
                        ->where('nomenclature', 'RX')
                        ->where('status', '0')
                        ->get();
                break;

                case 1:
                    $code_temp_last = CodePatient::where('patient_id', $id_temp)
                        ->where('nomenclature', 'RX')
                        ->where('status', '0')
                        ->get();
                break;

                case 2:
                    $code_temp_last = CodePatient::where('patient_id', $id_temp)
                        ->where('nomenclature', 'USG')
                        ->where('status', '0')
                        ->get();
                break;

                case 3:
                    $code_temp_last = CodePatient::where('patient_id', $id_temp)
                        ->where('nomenclature', 'MMO')
                        ->where('status', '0')
                        ->get();
                break;

                case 4:
                    $code_temp_last = CodePatient::where('patient_id', $id_temp)
                        ->where('nomenclature', 'DMO')
                        ->where('status', '0')
                        ->get();
                break;

            endswitch;

            if(count($code_temp_last) > 0):
                $code_last = $code_temp_last;
            else:
                $code_last = [];
            endif;

            $today = Carbon::now()->format('Y-m-d');

            if($exam == "1"):

                //$appoinment_last = Appointment::where('patient_id', $id_temp)
                    //->whereDate('date', '>=', $today)
                   // ->where('area','0')
                    //->orderBy('created_at', 'desc')
                    //->get();
                //$detalles = DetailAppointment::with(['service', 'study'])->get();
            else:
                //$appoinment_last = Appointment::where('patient_id', $id_temp)
                    /*->whereDate('date', '>=', $today)
                    ->where('area',$exam)
                    ->orderBy('created_at', 'desc')
                    ->get();*/
               // $detalles = DetailAppointment::with(['service', 'study'])->get();
            endif;       
    

           // if(count($appoinment_last) == 0):    
                /*if(count($code_last) == 0):
                    
                    $data = [
                        'patient' => $patient,
                        'exam' => $exam
                    ];
                else:
                    $data = [
                        'patient' => $patient,
                        'code_last' => $code_last,
                        'exam' => $exam
                    ];
                endif;  */
                
            //else:

                if(count($code_last) == 0):
                    $data = [
                        'patient' => $patient,
                        //'appointment_last' => $appoinment_last,
                        //'detalles' => $detalles,
                        'exam' => $exam
                    ];
                else:
                    
    
                    $data = [
                        'patient' => $patient,
                        'code_last' => $code_last,
                        //'appointment_last' => $appoinment_last,
                       // 'detalles' => $detalles,
                        'exam' => $exam
                    ];
                endif;
            //endif;
        endif;

        
        
        

        return response()->json($data);
    }

    public function getPatientBeneficiario($code, $exam){
            switch($exam):
                case 0:
                    $code_temp_last = CodePatient::where('patient_id', $code)
                        ->where('nomenclature', 'RX')
                        ->where('status', '0')
                        ->get();
                break;

                case 1:
                    $code_temp_last = CodePatient::where('patient_id', $code)
                        ->where('nomenclature', 'RX')
                        ->where('status', '0')
                        ->get();
                break;

                case 2:
                    $code_temp_last = CodePatient::where('patient_id', $code)
                        ->where('nomenclature', 'USG')
                        ->where('status', '0')
                        ->get();
                break;

                case 3:
                    $code_temp_last = CodePatient::where('patient_id', $code)
                        ->where('nomenclature', 'MMO')
                        ->where('status', '0')
                        ->get();
                break;

                case 4:
                    $code_temp_last = CodePatient::where('patient_id', $code)
                        ->where('nomenclature', 'DMO')
                        ->where('status', '0')
                        ->get();
                break;

            endswitch;
            $today = Carbon::now()->format('Y-m-d');

            if($exam == "1"):
               /* $appoinment_last = Appointment::where('patient_id', $code)
                    ->whereDate('date', '>=', $today)
                    ->where('area','0')
                    ->orderBy('created_at', 'desc')
                    ->get();
                $detalles = DetailAppointment::with(['service', 'study'])->get();*/
            else:
                /*$appoinment_last = Appointment::where('patient_id', $code)
                    ->whereDate('date', '>=', $today)
                    ->where('area',$exam)
                    ->orderBy('created_at', 'desc')
                    ->get();
                $detalles = DetailAppointment::with(['service', 'study'])->get();*/
            endif;       

            if(count($code_temp_last) > 0):
                $code_last = $code_temp_last;
            else:
                $code_last = [];
            endif;


            if(count($code_last) == 0):
                    
                $data = [
                    /*'appointment_last' => $appoinment_last,
                    'detalles' => $detalles,*/
                ];
            else:
                $data = [
                    /*'appointment_last' => $appoinment_last,
                    'detalles' => $detalles,*/
                    'code_last' => $code_last,
                ];
            endif;

        
        
        

        return response()->json($data);
    }

    public function getCodePatient($code){
        $date = Carbon::now();

        $codes_ant = CodePatient::where('nomenclature', $code)
                                ->where('year',$date->year)
                                ->count();
        $nomenclature = $code;

        if($nomenclature == 'RX' && $date->year == '2022'):
            $correlative = $codes_ant +'1';

        elseif($nomenclature == 'USG' && $date->year == '2022'):
            $correlative = $codes_ant +'3031';
        elseif($nomenclature == 'MMO' && $date->year == '2022'):
            $correlative = $codes_ant +'50';
        elseif($nomenclature == 'DMO' && $date->year == '2022'):
            $correlative = $codes_ant +'150';
        else:
            $correlative = $codes_ant +'1';
        endif;
        
        
        $year = $date->format('Y');
        $year_short = $date->format('y');
        
        if($correlative < 10):
            $code_new = $nomenclature.'0'.$correlative.'-'.$year_short;
        else:
            $code_new = $nomenclature.$correlative.'-'.$year_short;
        endif;       

        $data = [
            'nomenclature' => $nomenclature,
            'correlative' => $correlative,
            'year' => $year,
            'code' => $code_new
        ];

        return $data; 

        //return response()->json($data);
    }

    public function getStudies($type){
        $studies = Studie::where('type', $type)->get();


        return response()->json($studies);
    }

    public function getAppointments($date, $area){
        
        $citas = Appointment::where('date', $date)
                ->where('area',$area)
                ->count();
        //return $citas;
        return response()->json($citas);
    }

    public function getSchedule($date, $area){
        $cant_citas = Appointment::select(DB::raw('schedule_id, count(*) as total'))
                    ->where('date', $date)
                    ->where('area', $area)
                    ->groupBy('schedule_id')
                    ->get();       
        $control_citas = ControlAppointment::where('date', $date)->get();

        $data = [
            'cant_citas' => $cant_citas,
            'control_citas' => $control_citas
        ];

        return $data;

        //return $cant_citas;
        return response()->json($cant_citas);
    }

    public function getStudiesControlDate($date){
        $control_citas = ControlAppointment::where('date', $date)
                    ->get();

        //return $control_citas;
       
        return response()->json($control_citas);
    }

    public function getHolyDays($date){
        $consulta_dia_festivo = SettingHolyDays::where('holy_day', $date)->get();

        if(count($consulta_dia_festivo) == 1){
            $dia_festivo = 1;
        }else{
            $dia_festivo = 0;
        }

        //return $dia_festivo;
        return response()->json($dia_festivo);
    } 

    public function getScheduleChange(){
        $horarios = Schedule::all();

        //return $cant_citas;
        return response()->json($horarios);
    }

    public function getAppointmentsView(){
        
        $appointments = DB::table('appointments')
                ->join('patients', 'appointments.patient_id', '=', 'patients.id')
                ->join('schedule', 'appointments.schedule_id', '=', 'schedule.id')
                ->select(DB::raw('CONCAT(patients.lastname,  \', \' , patients.name) AS title'), DB::raw('CONCAT(appointments.date,  \' \' , schedule.hour_in) AS start'), DB::raw('CONCAT(appointments.date,  \' \' , schedule.hour_out) AS end'), DB::raw('appointments.area AS area'), DB::raw('schedule.type AS type'))
                ->get(); 
           
        
        return $appointments;
        $data = [
           'appointments' => $appointments
        ];

        return response()->json($data);
    }

    public function getAppointmentsViewRx(){
        
        $appointments = DB::table('appointments')
                ->join('patients', 'appointments.patient_id', '=', 'patients.id')
                ->join('schedule', 'appointments.schedule_id', '=', 'schedule.id')
                ->select(DB::raw('CONCAT(patients.lastname,  \', \' , patients.name) AS title'), DB::raw('CONCAT(appointments.date,  \' \' , schedule.hour_in) AS start'), DB::raw('CONCAT(appointments.date,  \' \' , schedule.hour_out) AS end'), DB::raw('appointments.area AS area'))
                ->where('appointments.area', 0)
                ->get();
           
        
        return $appointments;
        $data = [
           'appointments' => $appointments
        ];

        return response()->json($data);
    }

    public function getAppointmentsViewUmd(){
        
        $appointments = DB::table('appointments')
                ->join('patients', 'appointments.patient_id', '=', 'patients.id')
                ->join('schedule', 'appointments.schedule_id', '=', 'schedule.id')
                ->select(DB::raw('CONCAT(patients.lastname,  \', \' , patients.name) AS title'), DB::raw('CONCAT(appointments.date,  \' \' , schedule.hour_in) AS start'), DB::raw('CONCAT(appointments.date,  \' \' , schedule.hour_out) AS end'), DB::raw('appointments.area AS area'), DB::raw('schedule.type AS type'))
                ->whereIn('appointments.area', ['2', '3', '4'])  
                ->get();
           
        
        return $appointments;
        $data = [
           'appointments' => $appointments
        ];

        return response()->json($data);
    }

    public function getPruebaConsulta($mes, $year){
        
        $response = Http::get('http://10.11.0.45/api/medicosRayosX.php', [
            'month' => $mes,
            'year' => $year,
        ]);

        return $response->json();
    }

    public function getDiasFestivos( $year){
        
        $response = Http::get('http://10.11.0.45/api/asuetosOficiales.php', [
            'year' => $year,
        ]);
        
        return $response->json();
    }

    public function getServices(){
        $servicios = Service::where('parent_id', 2)
                ->where('parent_id', '<>', 0)
                ->get();

        //return $cant_citas;
        return response()->json($servicios);
    }

    public function pruebaConsultaApi(){
        $url = 'https://servicios.igssgt.org/WServices/wsComunicacionMedi/wsComunicacionMedi.asmx/ConsultarAfiliado';
        $usuario = 'wsComunicacion';
        $password = 'Igss.ws2021';
        $afiliado = '2817830711202';

        // Construcción del XML SOAP
        $soapRequest = 
        '<?xml version="1.0" encoding="utf-8"?>'.
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'.
  '<soap:Body>'.
    '<ConsultarAfiliado xmlns="http://tempuri.org/">'.
      '<Usuario>'.$usuario.'</Usuario>'.
      '<Clave>'.$password .'</Clave>'.
      '<numeroafiliado>'.$afiliado.'</numeroafiliado>'.
    '</ConsultarAfiliado>'.
  '</soap:Body>'.
'</soap:Envelope>';

        try {
           $response = Http::withHeaders([
                'Content-Type' => 'application/soap+xml; charset=utf-8',
                'Content-Length' => 'length',
                'SOAPAction' => 'http://tempuri.org/ConsultarAfiliado',
            ])->post(
                $url,
                $soapRequest
            );

            return response($response->body(), 200)->withHeaders([
                'Content-Type' => 'application/soap+xml; charset=utf-8',
                'Content-Length' => 'length'
            ]);
        } catch (\Exception $e) {
            return response()->json([
                'ok' => false,
                'mensaje' => 'Error al conectar con el servicio IGSS',
                'error' => $e->getMessage()
            ], 500);
        }

    }
    
}
