<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Models\User;
use App\Models\Appointment, App\Models\DetailAppointment, App\Models\Service, App\Models\Bitacora;
use DB, PDF, Auth, Carbon\Carbon;
use Maatwebsite\Excel\Facades\Excel;

use App\Exports\EstadisticasDMOExport, App\Exports\EstadisticasMAMOExport, App\Exports\EstadisticasUSGExport, App\Exports\EstadisticasRXExport;

class ReportController extends Controller
{
    public function getHome(){ 

        $tecnicos = User::select('id', DB::raw('CONCAT(name, \' \' , lastname) AS nombre'), 'ibm')
                        ->whereIn('role', [3])
                        ->orderBy('ibm', 'ASC')
                        ->get();
        //$tecnicos = User::pluck('name','id');

        //return $tecnicos;

        $data = [
            'tecnicos' => $tecnicos,
        ];


        return view('admin.statistics.home', $data);
    }

    public function postStaticsBetweenDates(Request $request){
        $date_in = $request->date_in;
        $date_out = $request->date_out;

        $citas_agen = Appointment::whereBetween('date', [$date_in, $date_out])
                            ->count();
        
        $citas_aten = Appointment::whereBetween('date', [$date_in, $date_out])
                            ->where('status', '3')
                            ->count();

        $tecnicos = User::select('id', DB::raw('CONCAT(name,  \' \' , lastname) AS nombre'), 'ibm')
                            ->whereIn('role', [2,3,5])
                            ->orderBy('ibm', 'ASC')
                            ->get();
                            
        $citas_tec_g_rx = DB::table('appointments')
                            ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                            ->whereBetween('date', [$date_in, $date_out])
                            ->where('area', '0')
                            ->groupBy('ibm_tecnico_1')
                            ->get();
        $citas_tec_g_usg = DB::table('appointments')
                            ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                            ->whereBetween('date', [$date_in, $date_out])
                            ->where('area', '2')
                            ->groupBy('ibm_tecnico_1')
                            ->get();
        $citas_tec_g_mmo = DB::table('appointments')
                            ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                            ->whereBetween('date', [$date_in, $date_out])
                            ->where('area', '3')
                            ->groupBy('ibm_tecnico_1')
                            ->get();
        $citas_tec_g_dmo = DB::table('appointments')
                            ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                            ->whereBetween('date', [$date_in, $date_out])
                            ->where('area', '4')
                            ->groupBy('ibm_tecnico_1')
                            ->get();
        $citas_tec_g_total = DB::table('appointments')
                            ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                            ->whereBetween('date', [$date_in, $date_out])
                            ->groupBy('ibm_tecnico_1')
                            ->get();

        //return $citas_tec_rx;
        
        $data = [
            'date_in' => $date_in,
            'date_out' => $date_out,
            'citas_agen' => $citas_agen,
            'citas_aten' => $citas_aten,
            'tecnicos' => $tecnicos,
            'citas_tec_g_rx' => $citas_tec_g_rx,
            'citas_tec_g_usg' => $citas_tec_g_usg,
            'citas_tec_g_mmo' => $citas_tec_g_mmo,
            'citas_tec_g_dmo' => $citas_tec_g_dmo,
            'citas_tec_g_total' => $citas_tec_g_total

        ];

        return view('admin.statistics.filtrer_date', $data);
    }

    public function getStaticsMonth(){
        $month =  Carbon::now()->format('m');
        $today = Carbon::now()->format('Y-m-d');

        $services = Service::select('id', DB::raw('CONCAT(name) AS nombre'))
                        ->where('type', '1')
                        ->get();

        $citas_x_ser = DB::table('appointments')
                        ->join('details_appointments', 'appointments.id', 'details_appointments.idappointment')
                        ->select(DB::raw('COUNT(DISTINCT details_appointments.idappointment) AS total_citas'), DB::raw('details_appointments.idservice AS id_service'))
                        ->whereMonth('appointments.date', $month)
                        ->groupBy('details_appointments.idservice')
                        ->get();

        

        $estudios_x_ser = DB::table('appointments')
                        ->join('details_appointments', 'appointments.id', 'details_appointments.idappointment')
                        ->select(DB::raw('COUNT(details_appointments.idstudy) AS total_estudios'), DB::raw('details_appointments.idservice AS id_service'))
                        ->whereMonth('appointments.date', $month)
                        ->groupBy('details_appointments.idservice')
                        ->get();
        
        /*$materiales_x_stu = DB::table('appointments')
                        ->join('details_appointments', 'appointments.id', 'details_appointments.idappointment')
                        ->join('materials_appointments', 'details_appointments.idappointment', 'materials_appointments.idappointment')
                        ->select(DB::raw('SUM(materials_appointments.amount) AS total_materiales'), DB::raw('details_appointments.idservice AS id_service'))
                        ->whereMonth('appointments.date', $month)
                        ->groupBy('details_appointments.idservice', 'materials_appointments.material')
                        ->get();*/

        $materiales_x_stu = DetailAppointment::with(['materials'])
        ->join('appointments', 'details_appointments.idappointment', 'appointments.id')
        ->whereMonth('appointments.date', $month)
        ->where('appointments.status', '3')
        ->get();



        return $materiales_x_stu;

        $data = [
            
            'today' => $today,
            'services' => $services,
            'citas_x_ser' => $citas_x_ser,
            'estudios_x_ser' => $estudios_x_ser
        ];

        //return $data;

        return view('admin.statistics.statistic_month',$data);
    }

    public function postReportRXEstadistica(Request $request){


        $mes = $request->get('month_rx');
        $year = $request->get('year_rx');

        // DEBUG: Verifica que estos valores sean lo que esperas (ej: "3" y "2026")
        // dd($mes, $year); 

        /*$datos = DB::table('details_appointments')
            ->select(
                DB::raw('Day(appointments.date) AS dia'),
                'services.id AS idservicio',
                DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes')
            )
            ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
            ->join('services', 'services.id', '=', 'details_appointments.idservice')
            ->whereMonth('appointments.date', $mes) // Usa directamente $mes
            ->whereYear('appointments.date', $year)
            ->where('appointments.status', 3)
            ->where('services.status', 1)
            ->groupBy('appointments.date', 'services.id')
            ->get();*/

        $datos = DB::table('details_appointments')
            ->select(
                '*'
            )
            ->where('details_appointments.idappointment','=',256701)

            ->get();


        return $datos;

        /*$mes = $request->get('month_rx');
        $month_in= getMonths(null, $mes);
        $year = $request->get('year_rx');

        $conteo_peliculas_11_14 = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('SUM(materials_appointments.amount) AS total_material'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'details_appointments.idappointment')
                    ->whereMonth('appointments.date', $mes)
                    ->whereYear('appointments.date', $year)
                    ->where('appointments.area', 0)
                    ->where('appointments.status', 3)
                    ->where('materials_appointments.material', 2)
                    ->groupBy(DB::raw('Day(appointments.date)'), 'materials_appointments.material')
                    ->get();
        
        return $conteo_peliculas_11_14 ;

        $b = new Bitacora;
        $b->action = "Generación de reporte mensual de RX del mes: ".$month_in.' - '.$year;
        $b->user_id = Auth::id();
        $b->save();

        $data = [
            'mes' => $mes,
            'year' => $year
        ];

        return Excel::download(new EstadisticasRXExport($data), 'Reporte RX '.$month_in.' - '.$year.'.xlsx');*/
    }

    public function postReportUSGEstadistica(Request $request){

        $mes = $request->get('month_usg');
        $month_in= getMonths(null, $mes);
        $year = $request->get('year_usg');

        /*$conteo_pacientes_hosp = DB::table('details_appointments')
            ->select(
                DB::raw('Day(appointments.date) AS dia'), 
                DB::raw('services.id AS idservicio'), 
                DB::raw('services.name AS servicio'),  
                DB::raw('COUNT(DISTINCT appointments.patient_id) AS total_pacientes'))
            ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
            ->join('services', 'services.id', '=', 'details_appointments.idservice')
            ->whereDay('appointments.date', 2)
            ->whereMonth('appointments.date', 2)
            ->whereYear('appointments.date', 2023)
            ->where('appointments.area', 2) 
            ->where('appointments.status', 3)
            ->where('services.parent_id', 1)
            ->groupBy(DB::raw('Day(appointments.date)'), DB::raw('services.id'))            
            ->get();
            
        return $conteo_pacientes_hosp;*/

        $b = new Bitacora;
        $b->action = "Generación de reporte mensual de USG del mes: ".$month_in.' - '.$year;
        $b->user_id = Auth::id();
        $b->save();
        
        $data = [
            'mes' => $mes,
            'year' => $year
        ];

        return Excel::download(new EstadisticasUSGExport($data), 'Reporte USG '.$month_in.' - '.$year.'.xlsx');
    }

    public function postReportMAMOEstadistica(Request $request){
        $mes = $request->get('month_mamo');
        $month_in= getMonths(null, $mes); 
        $year = $request->get('year_mamo');

        /*$conteo_peliculas_10_12 = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('SUM(materials_appointments.amount) AS total_material'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'details_appointments.idappointment')
                    ->whereMonth('appointments.date', $mes)
                    ->whereYear('appointments.date', $year)
                    ->where('appointments.area', 3)
                    ->where('appointments.status', 3)
                    ->where('materials_appointments.material', 0)
                    ->groupBy(DB::raw('Day(appointments.date)'), 'materials_appointments.material')
                    ->get();
        
        return $conteo_peliculas_10_12;*/

        $b = new Bitacora;
        $b->action = "Generación de reporte mensual de MMO del mes: ".$month_in.' - '.$year;
        $b->user_id = Auth::id();
        $b->save();

        $data = [
            'mes' => $mes,
            'year' => $year
        ];

        return Excel::download(new EstadisticasMAMOExport($data), 'Reporte MAMO '.$month_in.' - '.$year.'.xlsx');
    }

    public function postReportDMOEstadistica(Request $request){ 
        $mes = $request->get('month_dmo');
        $month_in = getMonths(null, $mes);
        $year = $request->get('year_dmo');

        /*$conteo_pacientes_coex = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('services.id AS idservicio'), 
                        DB::raw('services.name AS servicio'), 
                        DB::raw('COUNT(appointments.patient_id) AS total_pacientes'))
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('services', 'services.id', '=', 'details_appointments.idservice')
                    ->whereMonth('appointments.date', $mes)
                    ->whereYear('appointments.date', $year)
                    ->where('appointments.area', 4)
                    ->where('appointments.status', 3)
                    ->where('services.parent_id', 2)
                    ->groupBy(DB::raw('Day(appointments.date)'), DB::raw('services.id'))
                    ->get();

        return $conteo_pacientes_coex;*/
         
       /* $consulta_prueba = Service::where('parent_id', 2)->count();

        return $consulta_prueba;*/

        $b = new Bitacora;
        $b->action = "Generación de reporte mensual de DMO del mes: ".$month_in.' - '.$year;
        $b->user_id = Auth::id();
        $b->save();

        $data = [
            'mes' => $mes,
            'year' => $year
        ];

        return Excel::download(new EstadisticasDMOExport($data), 'Reporte DMO '.$month_in.' - '.$year.'.xlsx');
    }

    public function postReportTecnicoIndividual(Request $request){
        switch($request->area):
            case 0:
                $appointments = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('CONCAT(patients.name,  \' \' , patients.lastname) AS paciente'),
                        DB::raw('patients.affiliation AS afiliacion'),
                        DB::raw('studies.name AS estudio'),
                        DB::raw('materials_appointments.material AS material'),
                        DB::raw('materials_appointments.amount AS cantidad_material'),
                        DB::raw('appointments.num_study AS num_exp'))                        
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('patients', 'patients.id', '=', 'appointments.patient_id')
                    ->join('studies', 'studies.id', '=', 'details_appointments.idstudy')
                    ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'appointments.id')
                    ->whereMonth('appointments.date', $request->month_r_tec)
                    ->whereYear('appointments.date', $request->year_r_tec)
                    ->where('appointments.ibm_tecnico_1', $request->idTecnico)
                    ->where('appointments.area', 0)
                    ->orderBy('appointments.date')
                    ->distinct()
                    ->get();
            break;

            case 3:
                $appointments = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('CONCAT(patients.name,  \' \' , patients.lastname) AS paciente'),
                        DB::raw('patients.affiliation AS afiliacion'),
                        DB::raw('studies.name AS estudio'),
                        DB::raw('materials_appointments.material AS material'),
                        DB::raw('materials_appointments.amount AS cantidad_material'),
                        DB::raw('appointments.num_study AS num_exp'))                        
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('patients', 'patients.id', '=', 'appointments.patient_id')
                    ->join('studies', 'studies.id', '=', 'details_appointments.idstudy')
                    ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'appointments.id')
                    ->whereMonth('appointments.date', $request->month_r_tec)
                    ->whereYear('appointments.date', $request->year_r_tec)
                    ->where('appointments.ibm_tecnico_1', $request->idTecnico)
                    ->where('appointments.area', 3)
                    ->orderBy('appointments.date')
                    ->distinct()
                    ->get();
            break;

            case 4:
                $appointments = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('CONCAT(patients.name,  \' \' , patients.lastname) AS paciente'),
                        DB::raw('patients.affiliation AS afiliacion'),
                        DB::raw('studies.name AS estudio'),
                        DB::raw('materials_appointments.material AS material'),
                        DB::raw('materials_appointments.amount AS cantidad_material'),
                        DB::raw('appointments.num_study AS num_exp'))                        
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('patients', 'patients.id', '=', 'appointments.patient_id')
                    ->join('studies', 'studies.id', '=', 'details_appointments.idstudy')
                    ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'appointments.id')
                    ->whereMonth('appointments.date', $request->month_r_tec)
                    ->whereYear('appointments.date', $request->year_r_tec)
                    ->where('appointments.ibm_tecnico_1', $request->idTecnico)
                    ->where('appointments.area', 4)
                    ->orderBy('appointments.date')
                    ->distinct()
                    ->get();
            break;

            case 5:
                $appointments = DB::table('details_appointments')
                    ->select(
                        DB::raw('Day(appointments.date) AS dia'), 
                        DB::raw('CONCAT(patients.name,  \' \' , patients.lastname) AS paciente'),
                        DB::raw('patients.affiliation AS afiliacion'),
                        DB::raw('studies.name AS estudio'),
                        DB::raw('materials_appointments.material AS material'),
                        DB::raw('materials_appointments.amount AS cantidad_material'),
                        DB::raw('appointments.num_study AS num_exp'))                        
                    ->join('appointments', 'appointments.id', '=', 'details_appointments.idappointment')
                    ->join('patients', 'patients.id', '=', 'appointments.patient_id')
                    ->join('studies', 'studies.id', '=', 'details_appointments.idstudy')
                    ->join('materials_appointments', 'materials_appointments.idappointment', '=', 'appointments.id')
                    ->whereMonth('appointments.date', $request->month_r_tec)
                    ->whereYear('appointments.date', $request->year_r_tec)
                    ->where('appointments.ibm_tecnico_1', $request->idTecnico)
                    ->whereIn('appointments.area', [0,3,4])
                    ->orderBy('appointments.area', 'asc')
                    ->distinct()
                    ->get();
            break;
        endswitch;
        $tecnico = User::findorFail($request->idTecnico);
        $fecha = getMonths(null, $request->month_r_tec).'/'.$request->year_r_tec;

        $b = new Bitacora;
        $b->action = "Generación de reporte SPS-645 del técnico: ".$tecnico->name.' '.$tecnico->lastname.'del mes: '.$fecha;
        $b->user_id = Auth::id();
        $b->save();

        //return $request;
        //return $appointment;

        $data = [
            'appointments' => $appointments,
            'tecnico' => $tecnico,
            'fecha' => $fecha
        ];

        $pdf = PDF::loadView('admin.statistics.reporte_tecnico_individual',$data)->setPaper('legal', 'landscape');
        return $pdf->stream('SPS-645 '.$request->month_r_tec.'.pdf');

        
    }

    

    

    
}
