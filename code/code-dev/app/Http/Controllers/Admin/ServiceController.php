<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;

use App\Models\Service, App\Models\Unit, App\Models\Bitacora;
use Validator, Str, Config, Auth, Session, DB, Response;

class ServiceController extends Controller
{


    public function getHome(){

        $environment = Service::where('parent_id', '0')->orderby('id','Asc')->get();
        $units = Unit::all();
            

        $data = [
            'environment' => $environment,
            'units' => $units
        ];

        return view('admin.services.home',$data);
    }

    //Servicios Generales
    public function postServicesGeneralAdd(Request $request){
        $rules = [
    		'name' => 'required'
    	];
    	$messagess = [
    		'name.required' => 'Se requiere un nombre para el servicio general.'
    	];

        $validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', 'Se ha producido un error.')->with('typealert', 'danger');
        else: 
            $sg = new Service;
            $sg->name = e($request->input('name'));
            $sg->unit_id = e($request->input('unit_id'));
            $sg->parent_id = '0';
            $sg->type = '0';

            if($sg->save()):
                $b = new Bitacora;
                $b->action = "Registro de servicio general ".$sg->name;
                $b->user_id = Auth::id();
                $b->save();

                return back()->with('messages', '¡Servicio general creado y guardado con exito!.')
                    ->with('typealert', 'success');
    		endif;
        endif;
    }

    public function getServicesGeneralEdit($id){
        $servicesg = Service::find($id);

        $data = [
            'servicesg' => $servicesg
        ];

        return view('admin.services.edit_sg', $data);
    }

    public function postServicesGeneralEdit(Request $request, $id){
        $rules = [
    		'name' => 'required'
    	];
    	$messagess = [
    		'name.required' => 'Se requiere un nombre para el servicio general.'
    	];

        $validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', 'Se ha producido un error.')->with('typealert', 'danger');
        else: 
            $sg = Service::findOrFail($id);
            $sg->name = e($request->input('name'));

            if($sg->save()):
                $b = new Bitacora;
                $b->action = "Actualización de servicio general ".$sg->name;
                $b->user_id = Auth::id();
                $b->save();

                return redirect('/admin/services_g')->with('messages', '¡Servicio generla actualizado y guardado con exito!.')
                    ->with('typealert', 'success');
    		endif;
        endif;
    }

    //Servicios
    public function getServicesGeneralServices($id){
        $environment = Service::findOrFail($id);
        $services = Service::where('parent_id', $id)->get();
        $id = $id;

        $data = [
            'environment' => $environment,
            'services' => $services,
            'id' => $id
        ];

        return view('admin.services.services', $data);
    }

    public function postServicesGeneralServicesAdd(Request $request){
        $rules = [
    		'name' => 'required'
    	];
    	$messagess = [
    		'name.required' => 'Se requiere un nombre para el servicio general.'
    	];

        $validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', 'Se ha producido un error.')->with('typealert', 'danger');
        else: 
            $s = new Service;
            $s->name = e($request->input('name'));
            $s->parent_id = $request->input('parent_id');
            $s->type = '1'; 

            if($s->save()):
                $b = new Bitacora;
                $b->action = "Registro de servicio ".$s->name." del servicio general ".$s->service_parent->name;
                $b->user_id = Auth::id();
                $b->save();

                return back()->with('messages', '¡Servicio creado y guardado con exito!.')
                    ->with('typealert', 'success');
    		endif;
        endif;
    }
    
    public function getServicesGeneralServicesEdit($id){
        $environment = Service::findOrFail($id);
        $services = Service::find($id);
        $id = $id;

        $data = [
            'environment' => $environment,
            'services' => $services,
            'id' => $id
        ];

        return view('admin.services.edit_s', $data);
    }

    public function postServicesGeneralServicesEdit(Request $request, $id){
        $rules = [
    		'name' => 'required'
    	];
    	$messagess = [
    		'name.required' => 'Se requiere un nombre para el servicio general.'
    	];

        $validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', 'Se ha producido un error.')->with('typealert', 'danger');
        else: 
            $s = Service::findOrFail($id);
            $s->name = e($request->input('name'));

            if($s->save()):
                $b = new Bitacora;
                $b->action = "Actualización de datos de servicio ".$s->name." del servicio general ".$s->service_parent->name;
                $b->user_id = Auth::id();
                $b->save();

                return back()->with('messages', '¡Información del servicio actualizada y guardada con exito!.')
                    ->with('typealert', 'success');
    		endif;
        endif;
    }
    
    
}
