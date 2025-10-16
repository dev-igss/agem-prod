<?php

namespace App\Http\Controllers\Admin;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;

use App\Models\Setting, App\Models\Patient, App\Models\Unit, App\Models\CodePatient, App\Models\Appointment, App\Models\DetailAppointment, App\Models\Bitacora;

use App\Imports\importPatients;
use Validator, Str, Config, Auth, Session, DB, Response;

class PatientController extends Controller
{
    public function getHome($filtrado){

        switch ($filtrado) {

            case 'inicio':
                $patients = Patient::where('id',0)->get();  
            break;
    
            case 'a':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;

            case 'b':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'c':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;

            case 'ch':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'd':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'e':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'f':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'g':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'h':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'i':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'j':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'k':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'l':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'm':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'n':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'ñ':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'o':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'p':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'q':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'r':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 's':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 't':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'u':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'v':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'w':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'x':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'y':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;
    
            case 'z':
                $patients = Patient::with(['parent'])->where('lastname','LIKE',$filtrado.'%')->orderBy('lastname', 'Asc')->get();
            break;

            case 'borrados':
                $patients = Patient::onlyTrashed()->get();  
            break;
        }

                  
        //return $patients;
        $data = [
            'patients' => $patients
        ];

        return view('admin.patients.home',$data);
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

            if($request->input('type_search') == 1):
                $patients = Patient::with(['parent'])
                        ->where('affiliation','LIKE','%'.$request->input('search').'%')
                        ->orWhere('name','LIKE','%'.$request->input('search').'%')
                        ->orWhere('lastname','LIKE','%'.$request->input('search').'%')
                        ->get();
                $busqueda = $request->input('search');
                $tipo_busqueda = $request->input('type_search');
                $data = [
                    'patients' => $patients,
                    'busqueda' => $busqueda,
                    'tipo_busqueda' => $tipo_busqueda 
                ];
            else:
                $code = CodePatient::where('code',$request->input('search') )->first();

                if($code == NULL){
                    return redirect('admin/pacientes/inicio')->withErrors($validator)->with('messages', '¡No se encontro el código que busca!.')
                            ->with('typealert', 'danger')->withInput();
                }

                $patients = Patient::with(['parent'])
                        ->where('id',$code->patient_id)
                        ->get();
                $busqueda = $request->input('search');
                $tipo_busqueda = $request->input('type_search');
                $data = [
                    'patients' => $patients,
                    'busqueda' => $busqueda,
                    'tipo_busqueda' => $tipo_busqueda 
                ];
            endif;
            

            return view('admin.patients.search', $data); 
        
        endif;

        
    }

    public function getPatientAdd(){

        $units = Unit::all();
        

        $data = [
            'units' => $units
        ];

        return view('admin.patients.add', $data);
    }

    public function postPatientAdd(Request $request){
        $rules = [
            'affiliation' => 'required',
    		'name' => 'required',
            'lastname' => 'required'
    	];
    	$messagess = [
            'affiliation.required' => 'Se requiere el numero de afiliacion del paciente.',
    		'name.required' => 'Se requiere el nombre del paciente.',
            'lastname.required' => 'Se requiere los apellidos del paciente.'
    	];

        $validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', 'Se ha producido un error.')->with('typealert', 'danger');
        else: 
            $type_patient = $request->input('type_patient');
            $p = new Patient;
            $p->affiliation = e($request->input('affiliation'));
            if($type_patient == '1' || $type_patient == '2'):
                $patient_parent = Patient::where('affiliation', $request->input('af_prin'))->get();

                if(count($patient_parent) > 0):
                    foreach($patient_parent as $pp):
                        $idpatient_parent = $pp->id;
                    endforeach;
                    $p->affiliation_idparent = $idpatient_parent;
                else:
                    $p->affiliation_principal = $request->input('af_prin');
                endif;
            endif;
            $p->type = $request->input('type_patient');
            $p->name = e($request->input('name'));
            $p->lastname = e($request->input('lastname'));
            $p->unit_id = '1';
            //$p->exp_prev = e($request->input('exp_prev'));
            $p->age = $request->input('age');
            $p->birth = $request->input('birth'); 
            $p->gender = $request->input('gender');
            $p->contact = e($request->input('contact'));          

            if($p->save()):
                if($request->input('num_rx') != ""):
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;                                      
                    $cp->nomenclature = $request->input('num_rx_nom');
                    $cp->correlative = $request->input('num_rx_cor');
                    $cp->year = $request->input('num_rx_y');
                    $cp->code = $request->input('num_rx');
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_rx') == "" && $request->input('num_rx_nom') != "" && $request->input('num_rx_cor') != "" && $request->input('num_rx_y') != "" ):
                    $code_manual_rx = $request->input('num_rx_nom').$request->input('num_rx_cor').'-'.substr($request->input('num_rx_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;                                      
                    $cp->nomenclature = $request->input('num_rx_nom');
                    $cp->correlative = $request->input('num_rx_cor');
                    $cp->year = $request->input('num_rx_y');
                    $cp->code = $code_manual_rx;
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_usg') != ""):
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_usg_nom');
                    $cp->correlative = $request->input('num_usg_cor');
                    $cp->year = $request->input('num_usg_y');
                    $cp->code = $request->input('num_usg');
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_usg') == "" && $request->input('num_usg_nom') != "" && $request->input('num_usg_cor') != "" && $request->input('num_usg_y') != ""):
                    $code_manual_usg = $request->input('num_usg_nom').$request->input('num_usg_cor').'-'.substr($request->input('num_usg_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_usg_nom');
                    $cp->correlative = $request->input('num_usg_cor');
                    $cp->year = $request->input('num_usg_y');
                    $cp->code = $code_manual_usg;
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_mmo') != ""):
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_mmo_nom');
                    $cp->correlative = $request->input('num_mmo_cor');
                    $cp->year = $request->input('num_mmo_y');
                    $cp->code = $request->input('num_mmo');
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_mmo') == "" && $request->input('num_mmo_nom') != "" && $request->input('num_mmo_cor') != "" && $request->input('num_mmo_y') != ""):
                    $code_manual_mmo = $request->input('num_mmo_nom').$request->input('num_mmo_cor').'-'.substr($request->input('num_mmo_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_mmo_nom');
                    $cp->correlative = $request->input('num_mmo_cor');
                    $cp->year = $request->input('num_mmo_y');
                    $cp->code = $code_manual_mmo;
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_dmo') != ""):
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_dmo_nom');
                    $cp->correlative = $request->input('num_dmo_cor');
                    $cp->year = $request->input('num_dmo_y');
                    $cp->code = $request->input('num_dmo');
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_dmo') == "" && $request->input('num_dmo_nom') != "" && $request->input('num_dmo_cor') != "" && $request->input('num_dmo_y') != ""):
                    $code_manual_dmo = $request->input('num_dmo_nom').$request->input('num_dmo_cor').'-'.substr($request->input('num_dmo_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_dmo_nom');
                    $cp->correlative = $request->input('num_dmo_cor');
                    $cp->year = $request->input('num_dmo_y');
                    $cp->code = $code_manual_dmo;
                    $cp->status = '0';
                    $cp->save();
                endif;

                $b = new Bitacora;
                $b->action = "Registro de paciente con afiliacion: ".$p->affiliation;
                $b->user_id = Auth::id();
                $b->save();

                return redirect('/admin/pacientes/a')->with('messages', '¡Paciente creado y guardado con exito!.')
                    ->with('typealert', 'success');
    		endif;
        endif;
    }

    public function getPatientEdit($id){
        $patient = Patient::find($id);
        $code_rx = CodePatient::select('code')
                ->where('patient_id',$id)
                ->where('nomenclature', 'RX')
                ->where('status', '0')
                ->get();
        $code_usg = CodePatient::select('code')
            ->where('patient_id',$id)
            ->where('nomenclature', 'USG')
            ->where('status', '0')
            ->get();
        $code_mmo = CodePatient::select('code')
            ->where('patient_id',$id)
            ->where('nomenclature', 'MMO')
            ->where('status', '0')
            ->get();
        $code_dmo = CodePatient::select('code')
            ->where('patient_id',$id)
            ->where('nomenclature', 'DMO')
            ->where('status', '0')
            ->get();

            
        $data = [
            'patient' => $patient,
            'code_rx' => $code_rx,
            'code_usg' => $code_usg,
            'code_mmo' => $code_mmo,
            'code_dmo' => $code_dmo,
        ];

        return view('admin.patients.edit', $data);
    }

    public function postPatientEdit(Request $request, $id){

        $rules = [
    		'name' => 'required',
            'lastname' => 'required'
    	];
    	$messagess = [

    		'name.required' => 'Se requiere el nombre del paciente.',
            'lastname.required' => 'Se requiere los apellidos del paciente.'
    	];

        $validator = Validator::make($request->all(), $rules, $messagess);
    	if($validator->fails()):
    		return back()->withErrors($validator)->with('messages', 'Se ha producido un error.')->with('typealert', 'danger');
        else: 
            $p = Patient::findOrFail($id);
            $affiliation_ant = $p->affiliation;
            $type_patient = $request->input('type_patient');
            if($request->input('update_affiliation') != ""):
                if($affiliation_ant != $request->input('update_affiliation')):
                    $p->affiliation = $request->input('update_affiliation');
                else:
                    $p->affiliation = $affiliation_ant;
                endif;

                if($request->input('af_prin') != ""):
                    if($type_patient == '1' || $type_patient == '2'):
                        $patient_parent = Patient::where('affiliation', $request->input('af_prin'))->get();
        
                        if(count($patient_parent) > 0):
                            foreach($patient_parent as $pp):
                                $idpatient_parent = $pp->id;
                            endforeach;
                            $p->affiliation_idparent = $idpatient_parent;
                            $p->affiliation_principal = NULL;
                        else:
                            $p->affiliation_principal = $request->input('af_prin');
                            $p->affiliation_idparent = NULL;
                        endif;
                    endif;
                endif;
            else: 
                if($request->input('af_prin') != ""):
                    if($type_patient == '1' || $type_patient == '2'):
                        $patient_parent = Patient::where('affiliation', $request->input('af_prin'))->get();
        
                        if(count($patient_parent) > 0):
                            foreach($patient_parent as $pp):
                                $idpatient_parent = $pp->id;
                            endforeach;
                            $p->affiliation_idparent = $idpatient_parent;
                            $p->affiliation_principal = NULL;
                        else:
                            $p->affiliation_principal = $request->input('af_prin');
                            $p->affiliation_idparent = NULL;
                        endif;
                    endif;
                endif;
            endif;
            $p->type = $request->input('type_patient');
            $p->name = e($request->input('name'));
            $p->lastname = e($request->input('lastname'));
            $p->gender = $request->input('gender');
            $p->age = $request->input('age');
            $p->birth = $request->input('birth');
            $p->contact = e($request->input('contact'));            

            if($p->save()):

                if($request->input('num_rx') != ""):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'RX')
                        ->where('status', '0')    
                        ->update(['status' => 1]);

                    $cp_new = new CodePatient;
                    $cp_new->patient_id = $p->id;
                    $cp_new->nomenclature = $request->input('num_rx_nom');
                    $cp_new->correlative = $request->input('num_rx_cor');
                    $cp_new->year = $request->input('num_rx_y');
                    $cp_new->code = $request->input('num_rx');
                    $cp_new->status = '0';
                    $cp_new->save();
                endif;

                if($request->input('num_rx') == "" && $request->input('num_rx_nom') != "" && $request->input('num_rx_cor') != "" && $request->input('num_rx_y') != "" ):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'RX')
                        ->where('status', '0')    
                        ->update(['status' => 1]);

                    $code_manual_rx = $request->input('num_rx_nom').$request->input('num_rx_cor').'-'.substr($request->input('num_rx_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;                                      
                    $cp->nomenclature = $request->input('num_rx_nom');
                    $cp->correlative = $request->input('num_rx_cor');
                    $cp->year = $request->input('num_rx_y');
                    $cp->code = $code_manual_rx;
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_usg') != ""):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'USG')
                        ->where('status', '0')    
                        ->update(['status' => 1]);

                    $cp_new = new CodePatient;
                    $cp_new->patient_id = $p->id;
                    $cp_new->nomenclature = $request->input('num_usg_nom');
                    $cp_new->correlative = $request->input('num_usg_cor');
                    $cp_new->year = $request->input('num_usg_y');
                    $cp_new->code = $request->input('num_usg');
                    $cp_new->status = '0';
                    $cp_new->save();
                endif;

                if($request->input('num_usg') == "" && $request->input('num_usg_nom') != "" && $request->input('num_usg_cor') != "" && $request->input('num_usg_y') != ""):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'USG')
                        ->where('status', '0')    
                        ->update(['status' => 1]);

                    $code_manual_usg = $request->input('num_usg_nom').$request->input('num_usg_cor').'-'.substr($request->input('num_usg_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_usg_nom');
                    $cp->correlative = $request->input('num_usg_cor');
                    $cp->year = $request->input('num_usg_y');
                    $cp->code = $code_manual_usg;
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_mmo') != ""):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'MMO')
                        ->where('status', '0')    
                        ->update(['status' => 1]);

                    $cp_new = new CodePatient;
                    $cp_new->patient_id = $p->id;
                    $cp_new->nomenclature = $request->input('num_mmo_nom');
                    $cp_new->correlative = $request->input('num_mmo_cor');
                    $cp_new->year = $request->input('num_mmo_y');
                    $cp_new->code = $request->input('num_mmo');
                    $cp_new->status = '0';
                    $cp_new->save();
                endif;

                if($request->input('num_mmo') == "" && $request->input('num_mmo_nom') != "" && $request->input('num_mmo_cor') != "" && $request->input('num_mmo_y') != ""):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'MMO')
                        ->where('status', '0')    
                        ->update(['status' => 1]);

                    $code_manual_mmo = $request->input('num_mmo_nom').$request->input('num_mmo_cor').'-'.substr($request->input('num_mmo_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_mmo_nom');
                    $cp->correlative = $request->input('num_mmo_cor');
                    $cp->year = $request->input('num_mmo_y');
                    $cp->code = $code_manual_mmo;
                    $cp->status = '0';
                    $cp->save();
                endif;

                if($request->input('num_dmo') != ""):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'DMO')
                        ->where('status', '0')    
                        ->update(['status' => 1]);

                    $cp_new = new CodePatient;
                    $cp_new->patient_id = $p->id;
                    $cp_new->nomenclature = $request->input('num_dmo_nom');
                    $cp_new->correlative = $request->input('num_dmo_cor');
                    $cp_new->year = $request->input('num_dmo_y');
                    $cp_new->code = $request->input('num_dmo');
                    $cp_new->status = '0';
                    $cp_new->save();
                endif;

                if($request->input('num_dmo') == "" && $request->input('num_dmo_nom') != "" && $request->input('num_dmo_cor') != "" && $request->input('num_dmo_y') != ""):
                    $cp = CodePatient::where('patient_id', $p->id)
                        ->where('nomenclature', 'DMO')
                        ->where('status', '0')    
                        ->update(['status' => 1]);
                        
                    $code_manual_dmo = $request->input('num_dmo_nom').$request->input('num_dmo_cor').'-'.substr($request->input('num_dmo_y'), -2);
                    $cp = new CodePatient;
                    $cp->patient_id = $p->id;
                    $cp->nomenclature = $request->input('num_dmo_nom');
                    $cp->correlative = $request->input('num_dmo_cor');
                    $cp->year = $request->input('num_dmo_y');
                    $cp->code = $code_manual_dmo;
                    $cp->status = '0';
                    $cp->save();
                endif;

                $b = new Bitacora;
                if($request->input('update_affiliation') != ""):
                    $b->action = "Actualización de datos del paciente con afiliacion: ".$affiliation_ant." a número de afiliación: ".$request->input('update_affiliation');
                else:
                    $b->action = "Actualización de datos del paciente con afiliacion: ".$affiliation_ant;
                endif;                
                $b->user_id = Auth::id();
                $b->save();

                return back()->with('messages', '¡Información de paciente actualizada y guardada con exito!.')
                    ->with('typealert', 'success');
    		endif;
        endif;
    }

    public function getPatientUpdateParent($id, $af_principal){
        $p = Patient::findOrFail($id);

        $patient_parent = Patient::where('affiliation', $af_principal)->get();

    
        if(count($patient_parent) > 0):
            foreach($patient_parent as $pp):
                $idpatient_parent = $pp->id;
            endforeach;
            $p->affiliation_idparent = $idpatient_parent;
            $p->affiliation_principal = NULL;
        else:
            $p->affiliation_principal = $af_principal;
            $p->affiliation_idparent = NULL;
        endif;

        if($p->save()):
            $b = new Bitacora;
            $b->action = "Actualización de datos de la afiliación principal, del afiliado no.: ".$p->affiliation;                
            $b->user_id = Auth::id();
            $b->save();

            return back()->with('messages', '¡Información de afiliación principal actualizada y guardada con exito!.')
                ->with('typealert', 'success');
        endif;

    }

    public function getPatientHistoryExam($id){
        $appointments_rx = Appointment::where('patient_id', $id)
                            ->where('area', '0')
                            ->orderBy('date', 'Desc')
                            ->get();
        
        $appointments_usg = Appointment::where('patient_id', $id)
                            ->where('area', '2')
                            ->orderBy('date', 'Desc')
                            ->get();
    
        $appointments_mmo = Appointment::where('patient_id', $id)
                            ->where('area', '3')
                            ->orderBy('date', 'Desc')
                            ->get();
        
        $appointments_dmo = Appointment::where('patient_id', $id)
                            ->where('area', '4')
                            ->orderBy('date', 'Desc')
                            ->get();
            
        $data = [
            'appointments_rx' => $appointments_rx,
            'appointments_usg' => $appointments_usg,
            'appointments_mmo' => $appointments_mmo,
            'appointments_dmo' => $appointments_dmo
        ];

        return view('admin.patients.history_exam', $data);
    }

    public function getPatientHistoryCode($id){
        $codes_rx = CodePatient::where('patient_id', $id)
                            ->where('nomenclature', 'RX')
                            ->orderBy('created_at', 'desc')
                            ->get();
        
        $codes_usg = CodePatient::where('patient_id', $id)
                            ->where('nomenclature', 'USG')
                            ->orderBy('created_at', 'desc')
                            ->get();
    
        $codes_mmo = CodePatient::where('patient_id', $id)
                            ->where('nomenclature', 'MMO')
                            ->orderBy('created_at', 'desc')
                            ->get();
        
        $codes_dmo = CodePatient::where('patient_id', $id)
                            ->where('nomenclature', 'DMO')
                            ->orderBy('created_at', 'desc')
                            ->get();
            
        $data = [
            'codes_rx' => $codes_rx,
            'codes_usg' => $codes_usg,
            'codes_mmo' => $codes_mmo,
            'codes_dmo' => $codes_dmo
        ];

        return view('admin.patients.history_code', $data);
    }

    public function getPatientDelete($id){
        $patient = Patient::findOrFail($id);

        $citas = Appointment::where('patient_id', $id)->get();

        if(count($citas) > 0):
            return back()->with('messages', '¡No puede eliminar este paciente porque tiene citas registradas!.')
                        ->with('typealert', 'danger');
        else:
            if($patient->delete()):
                $b = new Bitacora;
                $b->action = "Se borro el registro del paciente con afiliación no. ".$patient->affiliation;
                $b->user_id = Auth::id();
                $b->save();
    
                return redirect('/admin/pacientes/inicio')->with('messages', '¡Paciente enviado a la papelera de reciclaje!.')
                        ->with('typealert', 'success');
            endif;

        endif;

        
    }

    public function getPatientRestore($id){
        $patient = Patient::findOrFail($id);
        return $patient;
        if($patient->restore()):
            $b = new Bitacora;
            $b->action = "Se restauro el registro del paciente con afiliación no. ".$patient->affiliation;
            $b->user_id = Auth::id();
            $b->save();

            return back()->with('messages', '¡Paciente restaurar de la papelera de reciclaje!.')
                    ->with('typealert', 'success');
        endif;
    }

    public function getConfigPatient(){
        $id = '1';
        $config = Setting::findOrFail($id);

        $data = [
            'config' => $config
        ];

        return view('admin.patients.setting', $data);
    }

    public function postConfigPatient(Request $request){
        $id = '1';
        $config_appo = Setting::findOrFail($id);
        $config_appo->correlative_rx = $request->input('correlative_rx');
        $config_appo->correlative_usg = $request->input('correlative_usg');
        $config_appo->correlative_mmo = $request->input('correlative_mmo');
        $config_appo->correlative_dmo = $request->input('correlative_dmo');

        if($config_appo->save()):
            $b = new Bitacora;
            $b->action = "Cambios en configuración de Correlativos";
            $b->user_id = Auth::id();
            $b->save();

            return back()->with('messages', '¡Configuraciones de correlativos actualizadas y guardadas con exito!.')
                ->with('typealert', 'success');
        endif;
    }


}
