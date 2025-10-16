<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

use App\Models\Patient, App\Models\CodePatient, App\Models\Appointment, App\Models\Service, App\Models\Studie;
use Carbon\Carbon;

class ApiController extends Controller
{
    
    public function getStudyName($id){
        $study = Studie::findOrFail($id);
        $name = $study->name;

        $data = [
            'name' => $name
        ];

        //return $data;
        return response()->json($data);
    }

    public function getStudy($type){
        $study = Studie::select('id', 'name')->where('type', $type)->get();
        

        $data = [
            'study' => $study
        ];

        //return $data;
        return response()->json($data);
    }
}
