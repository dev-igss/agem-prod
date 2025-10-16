<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;

use App\Models\Unit, App\Models\Bitacora, App\Models\Location;
use Validator, Str, Config, Auth;

class UnitsController extends Controller
{


    public function getHome(){
        $units = Unit::orderBy('name', 'Asc')->get();
        $locations = Location::orderBy('code', 'Asc')->get();

        $data = [
            'units' => $units,
            'locations' => $locations
        ];

    	return view('admin.units.home', $data);
    }

    public function postUnitAdd(Request $request){
    	$rules = [
    		'name' => 'required'
    	];
    	$messagess = [
    		'name.required' => 'Se requiere un nombre para la unidad.'
    	];

    	$validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', '¡Se ha producido un error!.')
                ->with('typealert', 'danger')->withInput();
        else:

            $u = new Unit;
            $u->name = e($request->input('name'));
            $u->code = $request->input('code');
            $u->location_id = $request->input('location_id');

            if($u->save()):
                $b = new Bitacora;
                $b->action = "Registro de unidad ".$u->name;
                $b->user_id = Auth::id();
                $b->save();

                return back()->with('messages', '¡Unidad registrada y guardada con exito!.')
                    ->with('typealert', 'success');
    		endif;
    	endif;
    }

    public function getUnitEdit($id){
        $unit = Unit::find($id);

        $data = [
            'unit' => $unit
        ];

        return view('admin.units.edit', $data);
    }

    public function postUnitEdit(Request $request, $id){
        $rules = [
    		'name' => 'required'
    	];
    	$messagess = [
    		'name.required' => 'Se requiere un nombre para la unidad.'
    	];

        $validator = Validator::make($request->all(), $rules, $messagess);
        if($validator->fails()):
            return back()->withErrors($validator)->with('messages', '¡Se ha producido un error!.')
                ->with('typealert', 'danger')->withInput();
        else:


            $u = Unit::find($id);
            $u->name = e($request->input('name'));
            $u->code = $request->input('code');

            if($u->save()):
                $b = new Bitacora;
                $b->action = "Actualización de datos de unidad ".$u->name;
                $b->user_id = Auth::id();
                $b->save();

                return back()->with('messages', '¡Información de unidad actualizada y guardada con exito!.')
                    ->with('typealert', 'success');
            endif;
        endif;
    }

    public function getUnitDelete($id){
        $u = Unit::find($id);
        if($u->delete()):

            $b = new Bitacora;
            $b->action = "Eliminación de unidad ".$u->name;
            $b->user_id = Auth::id();
            $b->save();

            return back()->with('messages', '¡Unidad eliminada con exito!.')
            ->with('typealert', 'success');
        endif;
    }
}
