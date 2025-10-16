<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;

use App\Http\Models\Setting;
use Validator, Str, Config, Auth, Session, DB, Response;

class SettingController extends Controller
{
    public function __Construct(){
        $this->middleware('auth');
        $this->middleware('IsAdmin');
        $this->middleware('UserStatus');
        $this->middleware('Permissions');
    }

    public function getConfigAppoint(){
        $config_appo = Setting::all();

        $data = [
            'config_appo' => $config_appo
        ];

        return view('admin.appointments.settings',$data);
    }

    public function postConfigAppo(Request $request,$id){
        $config_appo = Setting::findOrFail($id);
        $config_appo->amount_rx = "";
        $config_appo->amount_rx_emer_hosp = "";
        $config_appo->amount_rx_special_am = "";
        $config_appo->amount_rx_special_pm = "";
        $config_appo->amount_usg_emer_hosp = "";
        $config_appo->amount_usg_am = "";
        $config_appo->amount_usg_pm = "";
        $config_appo->amount_usg_doppler_am = "";
        $config_appo->amount_usg_doppler_pm = "";
        $config_appo->amount_mmo_emer_hosp = "";
        $config_appo->amount_mmo_am = "";
        $config_appo->amount_mmo_pm = "";
        $config_appo->amount_dmo_emer_hosp = "";
        $config_appo->amount_dmo_am = "";
        $config_appo->amount_dmo_pm = "";

        if($config_appo->save()):
            $b = new Bitacora;
            $b->action = "Cambios en configuración de Citas ";
            $b->user_id = Auth::id();
            $b->save();

            return back()->with('messages', '¡Configuraciones actualizadas y guardadas con exito!.')
                ->with('typealert', 'success');
        endif;
    }
}
