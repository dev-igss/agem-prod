<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Models\Schedule, App\Models\Bitacora;
use Validator, Str, Config, Auth, Session, DB, Response;


class ScheduleController extends Controller
{

    public function getHome(){

        $schedules = Schedule::all();            

        $data = [
            'schedules' => $schedules
        ];

        return view('admin.schedule.home', $data);
    }

    public function postScheduleAdd(Request $request){
    	$rules = [
    		'hour_in' => 'required',
            'hour_out' => 'required'
    	];
    	$messagess = [
    		'hour_in.required' => 'Se requiere la hora de inicio.',
            'hour_out.required' => 'Se requiere la hora de finalizacion.'
    	];

    	$validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', '¡Se ha producido un error!.')
                ->with('typealert', 'danger')->withInput();
        else:

            $s = new Schedule;
            $s->hour_in = $request->input('hour_in');
            $s->hour_out = $request->input('hour_out');
            $s->type = $request->input('type');

            if($s->save()):
                $b = new Bitacora;
                $b->action = "Registro de horario de".$s->hour_in.' a '.$s->hour_out;
                $b->user_id = Auth::id();
                $b->save();

                return back()->with('messages', '¡Horario registrado y guardado con exito!.')
                    ->with('typealert', 'success');
    		endif;
    	endif;
    }
}
