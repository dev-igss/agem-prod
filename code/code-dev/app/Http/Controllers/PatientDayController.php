<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

use App\Models\Appointment, App\Models\DetailAppointment, App\Models\MaterialAppointment, App\Models\Bitacora, App\Models\User;
use Carbon\Carbon,DateTime;

class PatientDayController extends Controller
{
    public function getPatientDayRx(){
        $today = Carbon::now()->format('Y-m-d');
        $date = Carbon::now()->format('d-m-Y');
        $date1 = new DateTime('now');
        $date1->modify('last day of this month');
        $appointments = Appointment::where('date',$today)
                    ->where('area', '0')
                    ->where('status', '1')
                    ->get();
        //$detalles = DetailAppointment::all();

        $data = [
            'appointments' => $appointments,
            'date' => $date,
            'today' => $today,
            'date1' => $date1
        ];

        return view('patients_day.home_rx', $data);
    }

    public function getPatientDayUmd($filtrado){
        $today = Carbon::now()->format('Y-m-d');
        $date = Carbon::now()->format('d-m-Y');
        $date1 = new DateTime('now');
        $date1->modify('last day of this month');

        switch($filtrado):
            case 'todas':
                $appointments = Appointment::where('date',$today)
                    ->whereIn('area', ['2', '3', '4'])                   
                    ->where('status', '1')
                    ->get();
            break;

            case 'usg':
                $appointments = Appointment::where('date',$today)
                    ->where('area', '2')                   
                    ->where('status', '1')
                    ->get();
            break;

            case 'mmo':
                $appointments = Appointment::where('date',$today)
                    ->where('area', '3')                   
                    ->where('status', '1')
                    ->get();
            break;

            case 'dmo':
                $appointments = Appointment::where('date',$today)
                    ->where('area', '4')                   
                    ->where('status', '1')
                    ->get();
            break;
        endswitch;
        
        //$detalles = DetailAppointment::all();

        $data = [
            'appointments' => $appointments,
            'date' => $date,
            'today' => $today,
            'date1' => $date1
        ];

        return view('patients_day.home_umd', $data);
    }

    public function getPatientDayOpen($filtrado){
        $today = Carbon::now()->format('Y-m-d');

        switch($filtrado):
            case 'todas':
                $appointments = Appointment::with('details')->where('status', '1')
                    ->where('date','<',$today)
                    ->whereYear('date', '2023') 
                    ->orderBy('date', 'ASC')
                    ->get();

                $total = $appointments->count();


            break;

            case 'rx':
                $appointments = Appointment::with('details')->where('area', '0')     
                    ->where('date','<',$today)       
                    ->whereYear('date', '2023')        
                    ->where('status', '1')
                    ->orderBy('date', 'ASC')
                    ->get();
                
                $total = $appointments->count();
            break;

            case 'usg':
                $appointments = Appointment::with('details')->where('area', '2') 
                    ->where('date','<',$today)     
                    ->whereYear('date', '2023')  
                    ->whereYear('date', '2023')             
                    ->where('status', '1')
                    ->orderBy('date', 'ASC')
                    ->get();
                
                $total = $appointments->count();
            break;

            case 'mmo':
                $appointments = Appointment::with('details')->where('area', '3')   
                    ->where('date','<',$today)   
                    ->whereYear('date', '2023')              
                    ->where('status', '1')
                    ->orderBy('date', 'ASC')
                    ->get();

                $total = $appointments->count();
            break;

            case 'dmo':
                $appointments = Appointment::with('details')->where('area', '4')  
                    ->where('date','<',$today)     
                    ->whereYear('date', '2023')             
                    ->where('status', '1')
                    ->orderBy('date', 'ASC')
                    ->get();

                $total = $appointments->count();
            break;
        endswitch;
        
        //$detalles = DetailAppointment::all();
       
        //$date = Carbon::now()->format('d-m-Y');
        $data = [
            'appointments' => $appointments,
            'total' => $total
        ];

        return view('patients_day.home_open', $data);
    }

    public function getMaterials($id){
        $idappointment = $id;
        $appointment = Appointment::findOrFail($id);
        $detalle = DetailAppointment::where('idappointment',$appointment->id)->get();

        $data = [
            'idappointment' => $idappointment,
            'appointment' => $appointment,
            'detalle' => $detalle
        ];

        return view('patients_day.materials', $data);
    }

    public function postMaterials(Request $request){
        $ibm_tec1 = $request->get('ibm1');
        $tecnico1 = User::select('id','name','lastname')->where('ibm', $ibm_tec1)->limit(1)->get();
        
        foreach($tecnico1 as $tec1):
            $id_tec1 = $tec1->id;
            $name_tec1 = $tec1->name.' '.$tec1->lastname;
        endforeach;

        if($request->get('ibm2') != ""):
            $ibm_tec2 = $request->get('ibm2');
            $tecnico2 = User::select('id','name','lastname')->where('ibm', $ibm_tec2)->limit(1)->get();
            
            foreach($tecnico2 as $tec2):
                $id_tec2 = $tec2->id;
                $name_tec2 = $tec2->name.' '.$tec2->lastname;
            endforeach;
        endif;

        $cita = Appointment::findOrFail($request->get('appointmentid'));
        $cita->status = '3';
        $hora = Carbon::now()->addMinutes(10)->format('H:i');
        $cita->check_out = $hora;
        $cita->ibm_tecnico_1 = $id_tec1;
        $area_temp = $cita->area;
        if($request->get('ibm2') != ""):
            $cita->ibm_tecnico_2 = $id_tec2;
        endif;
        $cita->save();

        $idappointment = $request->get('idappointment');
        $idstudy = $request->get('idstudy');
        $material = $request->get('material');
        $cantidad = $request->get('cantidad');
        $cont = 0;

        while($cont < count($idappointment)){
            $materials = new MaterialAppointment();
            $materials->idappointment = $idappointment[$cont];
            $materials->idstudy = $idstudy[$cont];
            $materials->material = $material[$cont];
            $materials->amount = $cantidad[$cont];
            $materials->save();
            $cont = $cont+1;
        }    

        if($materials->save()):     
            
            $b = new Bitacora;
            if($request->get('ibm2') != ""):
                $b->action = "Tecnicos ".$name_tec1." y ".$name_tec2." finalizarón atención de la cita no. ".$cita->id;
            else:
                $b->action = "Tecnico ".$name_tec1." finalizo atención de la cita no. ".$cita->id;
            endif;
            $b->user_id = $id_tec1;
            $b->save();

            if($area_temp == '0'):
                return redirect('/citas_del_dia_rx')->with('messages', '¡Registro de Materiales Exitos!.')
                    ->with('typealert', 'success');
            else:
                return redirect('/citas_del_dia_umd/todas')->with('messages', '¡Registro de Materiales Exitos!.')
                    ->with('typealert', 'success');
            endif;

        endif;
    } 

    public function getAppointmentComment($id, $text){
        $appointment = Appointment::findOrFail($id);
        $appointment->comment = $text;

        if($appointment->save()):     

            return back()->with('messages', '¡Comentario registrado y guardado correctamente!.')
                ->with('typealert', 'success');
            

        endif;
    }

    public function getAppointmentReschedule($id){
        $appointment = Appointment::findOrFail($id);
        $appointment->status = "5";

        if($appointment->save()):     
            

            if($appointment->area == '0'):
                return redirect('/citas_del_dia_rx')->with('messages', '¡Acción realizada con exito!.')
                    ->with('typealert', 'success');
            else:
                return redirect('/citas_del_dia_umd/todas')->with('messages', '¡Acción realizada con exito!.')
                    ->with('typealert', 'success');
            endif;

        endif;

        //return $appointment;
    }

    public function getAppointmentNot($id){
        $appointment = Appointment::findOrFail($id);
        $appointment->status = "6";

        if($appointment->save()):     
            

            if($appointment->area == '0'):
                return redirect('/citas_del_dia_rx')->with('messages', '¡Acción realizada con exito!.')
                    ->with('typealert', 'success');
            else:
                return redirect('/citas_del_dia_umd/todas')->with('messages', '¡Acción realizada con exito!.')
                    ->with('typealert', 'success');
            endif;

        endif;

        //return $appointment;
    }

    public function getAppointmentAddExamen($id, $area, $study, $comment){
        $detalle = new DetailAppointment();
        $detalle->idappointment=$id;
        switch($area):
            case '0':
                $detalle->idservice = '76';
            break;

            case '2':
                $detalle->idservice = '77';
            break;

            case '3':
                $detalle->idservice = '85';
            break;

            case '4':
                $detalle->idservice = '86';
            break;
        endswitch;
        $detalle->idstudy=$study;
        if($comment != "vacio"):
            $detalle->comment = $comment;
        endif;

        if($detalle->save()):     
            return back()->with('messages', '¡Registro de Estudio Exito!.')
                ->with('typealert', 'success');

        endif;
    }
}
