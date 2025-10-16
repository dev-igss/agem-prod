@extends('admin.master')
@section('title','Agregar Paciente')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/all') }}" class="nav-link"><i class="fas fa-columns"></i> Pacientes</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/add') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Agregar</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <form method="POST" action="{{ url('/admin/paciente/agregar') }}">
            @csrf
        <div class="row">
            <div class="col-md-4 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-cogs"></i><strong> Información Basica</strong></h2>
                    </div>

                    <div class="inside">   
                        <label for="name" ><strong><sup style="color: red;">(*)</sup> Numero de Afiliacion: </strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            
                            <select name="type_patient" id="patient_type" class="form-select col-2" aria-label="Default select example">
                                    @foreach(getTypePatient('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                            <input type="text" class="form-control" name="affiliation">
                        </div>

                        <div id="af_prin">
                            <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Afiliacion Principal: </strong></label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <input type="text" class="form-control" name="af_prin">
                            </div>
                        </div>                        

                        <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Nombre:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="name">
                        </div>

                        <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Apellidos:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="lastname">
                        </div>

                    </div>

                </div>
            </div>

            <div class="col-md-4 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-cogs"></i><strong> Información Adicional</strong></h2>
                    </div>

                    <div class="inside"> 
                        <label for="name" ><strong>Genero:</strong></label>
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
                            <input type="text" class="form-control" name="age">
                        </div>

                        <label for="name" class="mtop16"><strong>Fecha de Nacimiento:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="date" class="form-control" name="birth">
                        </div>

                        <label for="name" class="mtop16"><strong> Numero de Contacto:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="contact">
                        </div>
                    </div>

                </div>
            </div>

            <div class="col-md-4 d-flex"> 
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-qrcode"></i><strong> Correlativos de Expedientes</strong></h2>
                    </div>

                    <div class="inside">  
                        <div style="display:none;">
                            <label for="name" ><strong> ¿Cuenta con expedientes previos?:</strong></label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <select class="form-select col-2" aria-label="Default select example">
                                        
                                    <option value="0">No</option>
                                    <option value="1">Si</option>
                                </select>
                            </div>
                        </div>

                        <label for="ibm"><strong> Numero RX:</strong></label>                                              
                        <div id="div_generate_code_rx" style="display: display; margin-top: 10px;">
                            <div class="input-group">         
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <input type="text" class="form-control" name="num_rx" id="pnum_rx" readonly>
                                <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_rx" ><i class="fas fa-qrcode"></i> Generar</a>
                            </div>
                        </div>
                        
                        <div class="input-group"> 
                            <div id="div_manual_code_rx" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="RX" name="num_rx_nom" id="pnum_rx_nom" >
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_rx_cor" id="pnum_rx_cor" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong> Año:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_rx_y" id="pnum_rx_y" >
                                    </div>
                                </div>
                                <hr>
                            </div>
                            
                        </div>

                        <label for="ibm" class="mtop16"><strong> Numero USG:</strong></label>
                        <div id="div_generate_code_usg" style="display: display; margin-top: 10px;">
                            <div class="input-group">  
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <input type="text" class="form-control" name="num_usg" id="pnum_usg" readonly>
                                <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_usg" ><i class="fas fa-qrcode"></i> Generar</a>
                            </div>
                        </div>
                        
                        <div class="input-group"> 
                            <div id="div_manual_code_usg" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="USG" name="num_usg_nom" id="pnum_usg_nom" >
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_usg_cor" id="pnum_usg_cor" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong> Año:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_usg_y" id="pnum_usg_y" >
                                    </div>
                                </div>
                                <hr>
                                
                            </div> 
                        </div>

                        <label for="ibm" class="mtop16"><strong> Numero MMO:</strong></label>
                        <div id="div_generate_code_mmo" style="display: display; margin-top: 10px;">
                            <div class="input-group">  
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <input type="text" class="form-control" name="num_mmo" id="pnum_mmo" readonly>
                                <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_mmo" ><i class="fas fa-qrcode"></i> Generar</a>
                            </div>
                        </div>

                        <div class="input-group">    
                            <div id="div_manual_code_mmo" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="MMO" name="num_mmo_nom" id="pnum_mmo_nom" >
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_mmo_cor" id="pnum_mmo_cor" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong> Año:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_mmo_y" id="pnum_mmo_y" >
                                    </div>
                                </div>
                                <hr>
                                
                            </div> 
                        </div>

                        <label for="ibm" class="mtop16"><strong> Numero DMO:</strong></label>
                        <div id="div_generate_code_dmo" style="display: display; margin-top: 10px;">
                            <div class="input-group"> 
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <input type="text" class="form-control" name="num_dmo" id="pnum_dmo" readonly>
                                <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_dmo" ><i class="fas fa-qrcode"></i> Generar</a>
                            </div>
                        </div>

                        <div class="input-group">   
                            <div id="div_manual_code_dmo" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" value="DMO" name="num_dmo_nom" id="pnum_dmo_nom" >
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong> Correlativo:</strong></label>
                                        <input type="text" class="form-control" value="" name="num_dmo_cor" id="pnum_dmo_cor" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong> Año:</strong></label>
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
                        <button class="btn btn-success mt-4" type="submit">Guardar</button>
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

            /*var exp_prev = document.getElementById('exp_prev');
            var generate_rx = document.getElementById('div_generate_code_rx');
            var manual_rx = document.getElementById('div_manual_code_rx');
            var generate_usg = document.getElementById('div_generate_code_usg');
            var manual_usg = document.getElementById('div_manual_code_usg');
            var generate_mmo = document.getElementById('div_generate_code_mmo');
            var manual_mmo = document.getElementById('div_manual_code_mmo');
            var generate_dmo = document.getElementById('div_generate_code_dmo');
            var manual_dmo = document.getElementById('div_manual_code_dmo');
            generate_rx.style.display = "block";
            manual_rx.style.display = "none";
            generate_usg.style.display = "block";
            manual_usg.style.display = "none";
            generate_mmo.style.display = "block";
            manual_mmo.style.display = "none";
            generate_dmo.style.display = "block";
            manual_dmo.style.display = "none";

            $('#exp_prev').click(function(){
                if(exp_prev.value == 1){
                    manual_rx.style.display = "block"; 
                    generate_rx.style.display = "none";
                    manual_usg.style.display = "block"; 
                    generate_usg.style.display = "none";
                    manual_mmo.style.display = "block"; 
                    generate_mmo.style.display = "none";
                    manual_dmo.style.display = "block"; 
                    generate_dmo.style.display = "none";
                }else{
                    manual_rx.style.display = "none";
                    generate_rx.style.display = "block";
                    manual_usg.style.display = "none";
                    generate_usg.style.display = "block";
                    manual_mmo.style.display = "none";
                    generate_mmo.style.display = "block";
                    manual_dmo.style.display = "none";
                    generate_dmo.style.display = "block";
                }
            });*/


        });
    </script>

@endsection