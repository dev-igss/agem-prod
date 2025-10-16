<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Models\User;
use App\Models\Appointment, App\Models\DetailAppointment, App\Models\Service;
use DB, PDF, Auth;
use Carbon\Carbon;


class DashboardController extends Controller
{

    public function getDashboard(){
        $today = Carbon::now()->format('Y-m-d');
        $last_days = Carbon::now()->subDays(7)->format('Y-m-d');
        $month =  Carbon::now()->format('m');
        $year =  Carbon::now()->format('Y');

        $citas_d = Appointment::where('date', $today)
                        ->count();
        
        $citas_d_aten = Appointment::where('date', $today)
                        ->where('status','1')
                        ->count();
            
        $citas_d_aus = Appointment::where('date', $today)
                        ->where('status','2')
                        ->count();

        $citas_agen = Appointment::whereMonth('date', $month)
                        ->count();

        $citas_aten = Appointment::whereMonth('date', $month)
                        ->where('status','3')
                        ->count();
        
        $citas_aus = Appointment::whereMonth('date', $month)
                        ->where('status','2')
                        ->count();

        /*ESTADISTICAS DE USUARIOS(TECNICOS) LOGGEADOS*/
        $id_logeado = Auth::user()->id;
        $citas_d_tec = Appointment::where('date', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->count();
        $citas_ld_tec = Appointment::where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                  ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->count();
        $citas_m_tec = Appointment::whereMonth('date', $month)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->count();

        $citas_atendidas = Appointment::with(['patient','details','materials'])
                        ->whereMonth('date', $month)
                        ->whereYear('date', $year)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                  ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->get();

        /* ESTADISTICA DE CITAS POR TECNICO */
        $citas_d_tec_rx = Appointment::where('date', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '0')
                        ->count();
        $citas_ld_tec_rx = Appointment::where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                  ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '0')
                        ->count();
        $citas_m_tec_rx = Appointment::whereMonth('date', $month)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '0')
                        ->count();

        $citas_d_tec_usg = Appointment::where('date', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '2')
                        ->count();
        $citas_ld_tec_usg = Appointment::where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                  ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '2')
                        ->count();
        $citas_m_tec_usg = Appointment::whereMonth('date', $month)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '2')
                        ->count();


        $citas_d_tec_mmo = Appointment::where('date', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '3')
                        ->count();
        $citas_ld_tec_mmo = Appointment::where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                  ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '3')
                        ->count();
        $citas_m_tec_mmo = Appointment::whereMonth('date', $month)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '3')
                        ->count();     
                        
        $citas_d_tec_dmo = Appointment::where('date', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '4')
                        ->count();
        $citas_ld_tec_dmo = Appointment::where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                  ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '4')
                        ->count();
        $citas_m_tec_dmo = Appointment::whereMonth('date', $month)
                        ->where(function ($query) use ($id_logeado) {
                            $query->where('ibm_tecnico_1', $id_logeado)
                                ->orWhere('ibm_tecnico_2', $id_logeado);
                        })
                        ->where('area', '4')
                        ->count(); 

        /* ESTADISTICA DE CITAS POR TECNICO */
        $citas_x_tec = DB::table('appointments')
                        ->join('users', 'appointments.ibm_tecnico_1', 'users.id')
                        ->select(DB::raw('COUNT(appointments.ibm_tecnico_1) AS total_citas'), DB::raw('CONCAT(users.name,  \' \' , users.lastname) AS tecnico'), DB::raw('users.ibm AS ibm_tecnico'), DB::raw('appointments.ibm_tecnico_1 as id_tec_cita'))
                        ->whereMonth('appointments.date', $month)                        
                        ->where('appointments.status', 3)
                        ->groupBy('appointments.ibm_tecnico_1','users.name','users.lastname','users.ibm')
                        ->get();

        $tecnicos = User::select('id', DB::raw('CONCAT(name,  \' \' , lastname) AS nombre'), 'ibm')
                        ->whereIn('role', [3,5])
                        ->orderBy('ibm', 'ASC')
                        ->get();

        /* ESTADISTICA DE CITAS POR SERVICIO */
        $services = Service::select('id', DB::raw('CONCAT(name) AS nombre'))
                        ->where('type', '1')
                        ->get();

        $citas_x_ser = DB::table('appointments')
                        ->join('details_appointments', 'appointments.id', 'details_appointments.idappointment')
                        ->select(DB::raw('COUNT(details_appointments.idappointment) AS total_citas'), DB::raw('details_appointments.idservice AS id_service'))
                        ->whereMonth('appointments.date', $month)                        
                        ->where('appointments.status', 3)
                        ->groupBy('details_appointments.idservice')
                        ->get();


        /* ESTADISTICA DE TECNICOS POR AREA */
        $citas_d_tec_g_rx = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date', $today)
                        ->where('area', '0')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_ld_tec_g_rx = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where('area', '0')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_m_tec_g_rx = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->whereMonth('date', $month)
                        ->where('area', '0')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_d_tec_g_usg = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date', $today)
                        ->where('area', '2')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_ld_tec_g_usg = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where('area', '2')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_m_tec_g_usg = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->whereMonth('date', $month)
                        ->where('area', '2')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_d_tec_g_mmo = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date', $today)
                        ->where('area', '3')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_ld_tec_g_mmo = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where('area', '3')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_m_tec_g_mmo = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->whereMonth('date', $month)
                        ->where('area', '3')
                        ->groupBy('ibm_tecnico_1')
                        ->get();
        
        $citas_d_tec_g_dmo = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date', $today)
                        ->where('area', '4')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_ld_tec_g_dmo = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->where('date','>=', $last_days)
                        ->where('date','<=', $today)
                        ->where('area', '4')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        $citas_m_tec_g_dmo = DB::table('appointments')
                        ->select(DB::raw('ibm_tecnico_1 AS id_tecnico'), DB::raw('COUNT(ibm_tecnico_1) AS total_citas'))
                        ->whereMonth('date', $month)
                        ->where('area', '4')
                        ->groupBy('ibm_tecnico_1')
                        ->get();

        //return $citas_ld_tec_g_rx;

        $data = [
           'citas_d' => $citas_d,
           'citas_d_aten' => $citas_d_aten,
           'citas_d_aus' => $citas_d_aus,
           'citas_agen' => $citas_agen,
           'citas_aten' => $citas_aten,
           'citas_aus' => $citas_aus,
           'citas_d_tec' => $citas_d_tec,
           'citas_ld_tec' => $citas_ld_tec,
           'citas_m_tec' => $citas_m_tec,
           'citas_atendidas' => $citas_atendidas,
           'citas_d_tec_rx' => $citas_d_tec_rx,
           'citas_ld_tec_rx' => $citas_ld_tec_rx,
           'citas_m_tec_rx' => $citas_m_tec_rx,
           'citas_d_tec_usg' => $citas_d_tec_usg,
           'citas_ld_tec_usg' => $citas_ld_tec_usg,
           'citas_m_tec_usg' => $citas_m_tec_usg,
           'citas_d_tec_mmo' => $citas_d_tec_mmo,
           'citas_ld_tec_mmo' => $citas_ld_tec_mmo,
           'citas_m_tec_mmo' => $citas_m_tec_mmo,
           'citas_d_tec_dmo' => $citas_d_tec_dmo,
           'citas_ld_tec_dmo' => $citas_ld_tec_dmo,
           'citas_m_tec_dmo' => $citas_m_tec_dmo,
           'tecnicos' => $tecnicos,
           'citas_x_tec' => $citas_x_tec,
           'services' => $services,
           'citas_x_ser' => $citas_x_ser,
           'citas_d_tec_g_rx'=> $citas_d_tec_g_rx,
           'citas_ld_tec_g_rx' => $citas_ld_tec_g_rx,
           'citas_m_tec_g_rx' => $citas_m_tec_g_rx,
           'citas_d_tec_g_usg'=> $citas_d_tec_g_usg,
           'citas_ld_tec_g_usg' => $citas_ld_tec_g_usg,
           'citas_m_tec_g_usg' => $citas_m_tec_g_usg,
           'citas_d_tec_g_mmo'=> $citas_d_tec_g_mmo,
           'citas_ld_tec_g_mmo' => $citas_ld_tec_g_mmo,
           'citas_m_tec_g_mmo' => $citas_m_tec_g_mmo,
           'citas_d_tec_g_dmo'=> $citas_d_tec_g_dmo ,
           'citas_ld_tec_g_dmo' => $citas_ld_tec_g_dmo ,
           'citas_m_tec_g_dmo' => $citas_m_tec_g_dmo,
           'today' => $today
        ];

        return view('admin.dashboard',$data);
    }

    public function getStaticsDates(){
        $data = [
            
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
                            ->whereIn('role', [3,5])
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

}
