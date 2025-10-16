<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;

use App\Models\Bitacora;
use Validator, Carbon\Carbon;

class BitacoraController extends Controller
{

    public function getBitacora(){
        $today = Carbon::now()->format('Y-m-d');
        $bitacoras = Bitacora::with(['user'])
            ->orderBy('id', 'Desc')
            ->whereDate('created_at',$today )
            ->get();

        $data = [
            'bitacoras' => $bitacoras
        ];

        return view('admin.bitacoras.home',$data);
    }

    public function postSearch(Request $request){
        $rules = [
            'search' => 'required'
        ];

        $messages = [
            'search.required' => 'El campo consulta es requerido.'
        ];

        $validator = Validator::make($request->all(),$rules,$messages);

        if($validator->fails()):
            return back()->withErrors($validator)->with('messages', '¡Se ha producido un error!.')
            ->with('typealert', 'danger')->withInput();
        else:
            $bitacoras = Bitacora::with(['user'])
                        ->where('action','LIKE','%'.$request->input('search').'%')
                        ->get();
            $busqueda = $request->input('search');
            $data = [
                'bitacoras' => $bitacoras,
                'busqueda' => $busqueda
            ];

            return view('admin.bitacoras.search', $data); 
        
        endif;

        
    }
}
