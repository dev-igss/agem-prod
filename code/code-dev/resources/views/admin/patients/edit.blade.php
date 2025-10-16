@extends('admin.master')
@section('title','Editar Paciente')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/all') }}" class="nav-link"><i class="fas fa-columns"></i> Equipos</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/add') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Editar: {{ $patient->name }}</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
    <form method="POST" action="{{ url('/admin/paciente/'.$patient->id.'/editar') }}">
    @csrf

        <div class="row">
            <div class="col-md-4 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-cogs"></i><strong> Información Basica</strong></h2>
                    </div>

                    <div class="inside">   
                        <label for="name"><strong><sup style="color: red;">(*)</sup> Número de Afiliación: </strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control col-md-2" value="{{getTypePatient(null, $patient->type)}}" name="type_patient_edit" readonly>
                            <input type="text" class="form-control" value="{{$patient->affiliation}}" name="affiliation" readonly>               
                            @if(kvfj(Auth::user()->permissions, 'patient_edit_affiliation'))
                                <a href="#" class="btn btn-sm btn-primary " id="btn_update_affiliation" ><i class="fas fa-qrcode"></i> Actualizar</a>    
                            @endif         
                            <a href="#" class="btn btn-sm btn-warning " id="btn_patient_medi" ><i class="fas fa-qrcode"></i> Verificar Medi</a>               
                        </div>

                        <div id="div_update_affiliation" style="display: none; margin-top: 5px;">
                            <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Afiliación Nueva:</strong></label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <select name="type_patient"  id="patient_type"  class="form-select col-2" aria-label="Default select example">
                                    @foreach(getTypePatient('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                            <input type="text" class="form-control" name="update_affiliation">                                                      
                                
                            </div>

                            <div id="af_prin">
                            <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Afiliación Principal: </strong></label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <input type="text" class="form-control" name="af_prin">
                            </div>
                        </div>  
                        </div> 

                        <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Nombre:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" value="{{$patient->name}}" name="name">
                        </div>

                        <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Apellidos:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" value="{{$patient->lastname}}" name="lastname">

                        </div>

                        @if($patient->type == '1' || $patient->type == '2')
                            <hr>
                            <h4 style="margin-left: 75px;"> <strong>Datos de Afiliación Principal </strong> </h4>

                            @if($patient->affiliation_idparent != NULL)
                                <label for="name" class="mtop16"><strong><sup style="color: red;"></sup> Número de Afiliación: </strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{getTypePatient(null, $patient->parent->type)}}" name="type_patient_edit" readonly>
                                    <input type="text" class="form-control" value="{{$patient->parent->affiliation}}" name="affiliation" readonly>                                           
                                </div>

                                <label for="name" class="mtop16"><strong><sup style="color: red;"></sup> Nombre:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$patient->parent->name}}" name="name" readonly>
                                </div>

                                <label for="name" class="mtop16"><strong><sup style="color: red;"></sup> Apellidos:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$patient->parent->lastname}}" name="lastname" readonly>
                                </div>
                            @else
                                <label for="name" class="mtop16"><strong><sup style="color: red;"></sup> Número de Afiliación: </strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>

                                    <select name="type_patient_edit" class="form-select col-2" aria-label="Default select example">
                                    @foreach(getTypePatient('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                                <input type="text" class="form-control" value="{{$patient->affiliation_principal}}" name="affiliation" readonly>                                          
                                </div>
                            @endif
                        @endif
                    </div>

                </div>
            </div>

            <div class="col-md-4 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-cogs"></i><strong> Información Adicional</strong></h2>
                    </div>

                    <div class="inside"> 
                        <label for="name"><strong>Genero:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>
                            <select name="gender" class="form-select" aria-label="Default select example">
                                    @foreach(getGenderPatient('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                        </div>

                        <label for="name" class="mtop16"><strong>Edad:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" value="{{ $patient->age }}" name="age">
                        </div>

                        <label for="name" class="mtop16"><strong>Fecha de Nacimiento:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" value="{{ $patient->birth }}" name="birth">
                        </div>

                        <label for="name" class="mtop16"><strong>Numero de Contacto:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" value="{{ $patient->contact }}" name="contact">
                        </div>
                    </div>

                </div>
            </div>

            <div class="col-md-4 d-flex"> 
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-qrcode"></i><strong> Numeros de Expedientes</strong></h2>
                    </div>

                    <div class="inside"> 
                        
                        <label for="ibm" ><strong>Numero RX:</strong> @foreach($code_rx as $crx) {{ $crx->code }} @endforeach</label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="num_rx" id="pnum_rx" readonly>
                            <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_rx" ><i class="fas fa-qrcode"></i> Actualizar</a>
                            @if(\Carbon\Carbon::now()->year <= 2022)
                                <a href="#" class="btn btn-sm btn-warning " id="btn_manual_code_rx" ><i class="fas fa-qrcode"></i> Act. Manual</a>
                            @endif
                            
                            <div id="div_manual_code_rx" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="RX" name="num_rx_nom" id="pnum_rx_nom" >
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_rx_cor" id="pnum_rx_cor" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Año:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_rx_y" id="pnum_rx_y" >
                                    </div>
                                </div>
                                
                            </div> 
                        </div>

                        <label for="ibm" class="mtop16"><strong>Numero USG:</strong> @foreach($code_usg as $cusg) {{ $cusg->code }} @endforeach</label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="num_usg" id="pnum_usg" readonly>
                            <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_usg" ><i class="fas fa-qrcode"></i> Actualizar</a>
                            @if(\Carbon\Carbon::now()->year <= 2022)
                                <a href="#" class="btn btn-sm btn-warning " id="btn_manual_code_usg" ><i class="fas fa-qrcode"></i> Act. Manual</a>
                            @endif
                            
                            <div id="div_manual_code_usg" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="USG" name="num_usg_nom" id="pnum_usg_nom" readonly>
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_usg_cor" id="pnum_usg_cor" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Año:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_usg_y" id="pnum_usg_y" readonly>
                                    </div>
                                </div>
                                
                            </div> 
                        </div>

                        <label for="ibm" class="mtop16"><strong>Numero MMO:</strong> @foreach($code_mmo as $cmmo) {{ $cmmo->code }} @endforeach</label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="num_mmo" id="pnum_mmo" readonly>
                            <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_mmo" ><i class="fas fa-qrcode"></i> Actualizar</a>
                            @if(\Carbon\Carbon::now()->year <= 2022)
                                <a href="#" class="btn btn-sm btn-warning " id="btn_manual_code_mmo" ><i class="fas fa-qrcode"></i> Act. Manual</a>
                            @endif
                            
                            <div id="div_manual_code_mmo" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="MMO" name="num_mmo_nom" id="pnum_mmo_nom" readonly>
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_mmo_cor" id="pnum_mmo_cor">
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Año:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_mmo_y" id="pnum_mmo_y">
                                    </div>
                                </div>
                                
                            </div> 
                        </div>

                        <label for="ibm" class="mtop16"><strong>Numero DMO:</strong> @foreach($code_dmo as $cdmo) {{ $cdmo->code }} @endforeach</label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="num_dmo" id="pnum_dmo" readonly>
                            <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_dmo" ><i class="fas fa-qrcode"></i> Actualizar</a>
                            @if(\Carbon\Carbon::now()->year <= 2022)
                                <a href="#" class="btn btn-sm btn-warning " id="btn_manual_code_dmo" ><i class="fas fa-qrcode"></i> Act. Manual</a>
                            @endif
                            
                            <div id="div_manual_code_dmo" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="DMO" name="num_dmo_nom" id="pnum_dmo_nom" readonly>
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_dmo_cor" id="pnum_dmo_cor" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong><sup style="color: red;">(*)</sup> Año:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_dmo_y" id="pnum_dmo_y" >
                                    </div>
                                </div>
                                
                            </div> 
                        </div>

                        
                    </div>

                </div>
            </div>
        </div>

        <div class="row mtop16">
            <div class="col-md-12">
                <div class="panel shadow">
                    <div class="inside">
                        <button class="btn btn-success mt-4" type="submit">Actualizar</button>
                    </div>
                </div>                    
            </div>
        </div>
        </form>
    </div>


    <script>
        $(document).ready(function(){
            
            var patient_type = document.getElementById('patient_type');
            var af_prin = document.getElementById('af_prin');
            
            af_prin.hidden = true;

            $('#patient_type').click(function(){
                if(patient_type.value == 1 || patient_type.value == 2){
                    af_prin.hidden = false;
                }else{
                    af_prin.hidden = true;
                }
            });
        });
    </script>
@endsection