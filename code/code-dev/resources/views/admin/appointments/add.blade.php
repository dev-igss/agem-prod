@extends('admin.master')
@section('title','Agendar Cita')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/all') }}" class="nav-link"><i class="fas fa-columns"></i> Citas</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/add') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Agendar Nueva</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        

        <form method="POST" action="{{ url('/admin/cita/agregar') }}">
        @csrf
        <div class="row mtop16">
            <div class="col-md-4 d-flex">
                <div class="panel shadow">
                    <!-- Modal -->

                    <div class="header">
                        <h2 class="title"><i class="fas fa-cogs"></i><strong> Información Paciente</strong></h2>
                    </div>

                    <div class="inside">   

                        <input type="hidden" class="form-control" name="patient_id" id="patient_id">
                        <input type="hidden" class="form-control" name="type_exam" id="type_examp">

                        <label for="name" ><strong> Numero de Afiliacion: </strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="affiliationp" id="affiliationp">
                            <select name="type" id="exam_b" class="form-select col-md-4" aria-label="Default select example">
                                    @foreach(getExamB('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                            <a href="#" class="btn btn-sm btn-info " id="btn_add_patient_search" data-toogle="tooltrip" data-placement="top" title="Buscar" ><i class="fas fa-search"></i> </a>
                            <a href="{{ url('/admin/cita/agregar') }}" class="btn btn-sm btn-warning " data-toogle="tooltrip" data-placement="top" title="Limpiar" ><i class="fas fa-sync-alt"></i> </a>
                        </div>

                        <label class="mtop16" style="color: red; font-size: 1em; margin-left: 50px; font-weight: bold; display: none;" id="patient_msg" >¡No se encontró al paciente, regístrelo por favor!</label>
                        
                        

                        <div class="input-group">

                            <div id="div_beneficiarios" style="display: none; margin-top: 10px; width: 100%;">
                                <label for="name" class="mtop16"><strong>¿Seleccionar Beneficiario?:</strong></label>
                                <div class="input-group ">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>

                                    <select name="beneficiario_question" id="beneficiario_question" style="width: 90%;">                                
                                            <option value="0">No</option>
                                            <option value="1">Sí</option>
                                    </select>
                                </div>

                                <div id="div_select_beneficiarios" style="display: none; margin-top: 10px; width: 100%;">
                                    <select name="beneficiarios" id="beneficiarios" class="form-select " aria-label="Default select example">
                                        <option selected>Seleccione una opción</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                            

                        <label for="name" class="mtop16"><strong> Nombre:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="namep" id="namep" readonly>
                        </div>

                        <label for="name" class="mtop16"><strong> Apellidos:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="lastnamep" id="lastnamep" readonly>
                        </div>

                        <label for="name" class="mtop16"><strong> Contacto:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="contactp" id="contactp" >
                        </div>

                        <label for="name" class="mtop16"><strong> Numero de Expediente Actual:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="numexpp" id="numexpp" readonly>
                            <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code_patient_actual" ><i class="fas fa-qrcode"></i> Generar</a>
                            @if(\Carbon\Carbon::now()->year <= 2022)
                                <a href="#" class="btn btn-sm btn-warning " id="btn_manual_code_patient_actual" ><i class="fas fa-qrcode"></i> Manual</a>
                            @endif
                        </div>

                        <div class="input-group"> 
                            <div id="div_manual_code_patient_actual" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" name="num_code_nom_act" id="num_code_nom_act" >
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong> Correlativo:</strong></label>
                                        <input type="text" class="form-control" name="num_code_cor_act" id="num_code_cor_act" >
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong> Año:</strong></label>
                                        <input type="text" class="form-control" name="num_code_y_act" id="num_code_y_act" >
                                    </div>
                                </div>
                            </div>
                            
                        </div>

                        

                        
                        

                    </div>

                </div>
            </div>

            <div class="col-md-4" id="register" style="display: none;"> 
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-calendar-alt"></i><strong> Registro de Paciente</strong></h2>
                    </div>

                    <div class="inside"> 

                        <label for="name"><strong> Nombre:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="name_new" >
                            <select name="type_patient_new" id="patient_type" class="form-select col-md-2" aria-label="Default select example">
                                    @foreach(getTypePatient('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                        </div>

                        <label for="name" class="mtop16"><strong> Apellidos:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="lastname_new">
                        </div>

                        <div class="row">

                            <div class="col-md-6">
                                <label for="name" class="mtop16"><strong>Genero:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>
                                    <select name="gender_new" class="form-select col-2" aria-label="Default select example">
                                    @foreach(getGenderPatient('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                                </div>
                            </div>

                            <div class="col-md-6">

                                <label for="name" class="mtop16"><strong>Edad:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" name="age_new">
                                </div>
                            </div>
                        </div>

                        <label for="name" class="mtop16"><strong> Contacto:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="contact_new">
                        </div>

                        <label for="name" class="mtop16"><strong> Numero de Expediente:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="num_code_new" id="num_code_new" readonly>
                            <a href="#" class="btn btn-sm btn-primary " id="btn_generate_code" ><i class="fas fa-qrcode"></i> Generar</a>
                            @if(\Carbon\Carbon::now()->year <= 2022)
                                <a href="#" class="btn btn-sm btn-warning " id="btn_manual_code" ><i class="fas fa-qrcode"></i> Manual</a>
                            @endif
                        </div>

                        <div class="input-group"> 
                            <div id="div_manual_code_new" style="display: none; margin-top: 10px;">
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="ibm" class="mtop16"><strong> Nomenclatura:</strong></label>
                                        <input type="text" class="form-control" name="num_code_nom" id="num_code_nom">
                                    </div>

                                    <div class="col-md-6"> 
                                        <label for="ibm" class="mtop16"><strong> Correlativo:</strong></label>
                                        <input type="text" class="form-control" name="num_code_cor" id="num_code_cor">
                                    </div>                                    
                                </div>

                                <div class="row">
                                    <div class="col-md-12">  
                                        <label for="ibm" class="mtop16"><strong> Año:</strong></label>
                                        <input type="text" class="form-control" name="num_code_y" id="num_code_y">
                                    </div>
                                </div>
                            </div>
                            
                        </div>

                        
                     
                        
                    </div>

                </div>
            </div>

      
            <div class="col-md-4">
                <div class="panel shadow">

                    <div class="header">
                        <h2 class="title"><i class="fas fa-calendar-plus"></i><strong> Información Para Cita Nueva</strong></h2>
                    </div>

                    <div class="inside"> 
                        
                        <label for="name"><strong>Fecha a Agendar:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="date" class="form-control" name="date" id="date_new_app" required>
                            <!--<a href="#" class="btn btn-sm btn-info " id="bt_ver_citas" ><i class="fa-solid fa-calendar-day"></i> Ver C.C</a>-->
                        </div>

                        <div id="dia-festivo" class="row mtop16" style="text-align:center; color:red; display: none;">
                            <div class="col-md-12 d-flex">
                                <div class="panel shadow"> 
                                    <!-- Modal -->

                                    <div class="header" >
                                        <h2 class="title"><i class="fa-solid fa-triangle-exclamation"></i><strong> Advertencia</strong></h2>
                                    </div>

                                    <div class="inside" style=" font-size: 8px;">
                                        <h3>¡Día festivo, solo puede registrar estudios para las áreas de emergencia y hospitalización!</h3>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div id="alert-control-citas" class="row mtop16" style="display: none;">
                            <div class="col-md-12 d-flex">
                                <div class="panel shadow">
                                    <!-- Modal -->

                                    <div class="header" >
                                        <h2 class="title"><i class="fas fa-calendar-day"></i><strong> Listado de Citas Agendadas Este Día</strong></h2>
                                    </div>

                                    <div class="inside">
                                        <div id="citas_agendadas_rx" style="display: none;">
                                            <label for="name" ><strong> Control de Citas RX:</strong></label> 
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1">AM</span>
                                                <input type="text" class="form-control col-md-4" name="control_rx_am" id="control_rx_am" rows="1" readonly>
                                                &emsp;&emsp;&emsp;&emsp;&emsp;
                                                <span class="input-group-text" id="basic-addon1">PM</span>
                                                <input type="text" class="form-control col-md-4" name="control_rx_pm" id="control_rx_pm" rows="1" readonly>
                                            </div>
                                        </div>

                                        <div id="citas_agendadas_usg" style="display: none;">
                                            <label for="name" ><strong> Control de Citas USG:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1">AM</span>
                                                <input type="text" class="form-control col-md-4" name="control_usg_am" id="control_usg_am" rows="1" readonly>
                                                &emsp;&emsp;&emsp;&emsp;&emsp;
                                                <span class="input-group-text" id="basic-addon1">PM</span>
                                                <input type="text" class="form-control col-md-4" name="control_usg_pm" id="control_usg_pm" rows="1" readonly>
                                            </div>

                                            <label for="name" class="mtop16"><strong> Control de Citas USG Doppler:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1">AM</span>
                                                <input type="text" class="form-control col-md-4" name="control_usg_doppler_am" id="control_usg_doppler_am" rows="1" readonly>
                                                &emsp;&emsp;&emsp;&emsp;&emsp;
                                                <span class="input-group-text" id="basic-addon1">PM</span>
                                                <input type="text" class="form-control col-md-4" name="control_usg_doppler_pm" id="control_usg_doppler_pm" rows="1" readonly>
                                            </div>
                                        </div>

                                        <div id="citas_agendadas_mmo" style="display: none;">
                                            <label for="name" ><strong> Control de Citas MMO:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1">AM</span>
                                                <input type="text" class="form-control col-md-4" name="control_mmo_am" id="control_mmo_am" rows="1" readonly>
                                                &emsp;&emsp;&emsp;&emsp;&emsp;
                                                <span class="input-group-text" id="basic-addon1">PM</span>
                                                <input type="text" class="form-control col-md-4" name="control_mmo_pm" id="control_mmo_pm" rows="1" readonly>
                                            </div>
                                        </div>

                                        <div id="citas_agendadas_dmo" style="display: none;">
                                            <label for="name" ><strong> Control de Citas DMO:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1">AM</span>
                                                <input type="text" class="form-control col-md-4" name="control_dmo_am" id="control_dmo_am" rows="1" readonly>
                                                &emsp;&emsp;&emsp;&emsp;&emsp;
                                                <span class="input-group-text" id="basic-addon1">PM</span>
                                                <input type="text" class="form-control col-md-4" name="control_dmo_pm" id="control_dmo_pm" rows="1" readonly>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <input type="hidden" class="form-control col-md-4" value="{{$setting_x_hour}}" name="citas_configuradas" id="citas_configuradas" >

                        <label for="name" class="mtop16"><strong>Horarios de Atención:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <select name="schedule" id="schedules" style="width: 88%" >
                                @foreach($schedules as $s)
                                    <option value=""></option>
                                    <option value="{{ $s->id }}">{{ \Carbon\Carbon::parse($s->hour_in)->format('H:i').' '.getHourType(null, $s->type) }}</option>
                                @endforeach
                            </select>
                        </div>

                        <label for="idsupplier" class="mtop16"><strong> Servicio Solicitante:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>                            
                            <select name="pidservice" id="pidservice" style="width: 88%" >
                                @foreach ($services as $s)
                                    <option value=""></option>
                                    <option value="{{ $s->id }}">{{ $s->name }}</option>
                                @endforeach
                            </select>
                        </div>

                        <label for="idsupplier" class="mtop16"><strong> Estudio / Examen A Realizar:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>
                            
                                <select name="studies" id="studies" class="form-select" style="width: 90%" aria-label="Default select example">
                                        <option></option>
                                        @foreach ($studies as $s)                                    
                                            <option value="{{$s->id}}">{{$s->name}}</option>
                                        @endforeach
                                    </select>
                            <input type="hidden" class="form-control col-md-4" value="0" name="studies_actual" id="studies_actual" >
                        </div>

                        <label for="name" class="mtop16"><strong> Comentario:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="textarea" class="form-control" name="pcomment" id="pcomment" rows="1">
                        </div>

                        <label for="name" class="mtop16"><strong>¿Paciente Presente?:</strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <select name="present_patient" class="form-select col-2" aria-label="Default select example">                                
                                <option value="0">No</option>
                                <option value="1">Sí</option>
                            </select>
                        </div>

                        <div class="input-group mtop16">
                            <a class="btn btn-primary mt-4" style="color: white;" id="btn_agregar" >Agregar</a>
                        </div>
                    </div>
                </div>
            </div>


            
        </div>

                
        <div class="row mtop16">
            

            <div class="col-md-12">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-calendar-plus"></i><strong> Detalle de Estudio(s) a Realizar</strong></h2>
                    </div>

                    <div class="inside">

                        <div class="card-body table-responsive">
                            <table id="detalles" class= "table table-striped table-bordered table-condensed table-hover">
                                <thead style="background-color: #c3f3ea">
                                    <th><strong> ELIMINAR </strong></th>
                                    <th><strong> SERVICIO </strong></th>
                                    <th><strong> ESTUDIO </strong></th>
                                    <th><strong> COMENTARIO </strong></th>
                                </thead>

                                <tbody>

                                </tbody>
                            </table>
                        </div>

                    </div>
                </div>          
                
                <div class="panel shadow mtop16">
                    <div class="inside">
                        <button class="btn btn-success mt-4" id="btn_guardar" type="submit">Guardar</button>
                    </div>
                </div>                 
            </div>
        </div>

        
        </form>
    </div>


    <script>
       
        var modal = document.getElementById('modelId');
        var modalVerCitas = document.getElementById('modalVerCitas');
        $("#btn_guardar").hide();
        var cont = 0;


        $(document).ready(function(){
            var today = new Date();
            var dd = today.getDate();
            var mm = today.getMonth() + 1; //January is 0!
            var yyyy = today.getFullYear();

            if (dd < 10) {
                dd = '0' + dd;
            }

            if (mm < 10) {
                mm = '0' + mm;
            } 
                
            today = yyyy + '-' + mm + '-' + dd;
            document.getElementById("date_new_app").setAttribute("min", today);




            $('#btn_agregar').click(function(){
                agregar_tabla();
            });

            $('#bt_add').click(function(){
                agregar();
            });

            

            $('#bt_search').click(function(){
                $('#modelId').modal("show");
            });
            $('#bt_closeModal').click(function(){
                $('#modelId').modal("hide");
            });

            $('#bt_ver_citas').click(function(){
                $('#modalVerCitas').modal("show");
            });

            $('#bt_cerrar_ver_citas').click(function(){
                $('#modalVerCitas').modal("hide");
            });
        });



        function limpiar(){
            $("#ppatient_id").val("");
            $("#affiliation_b").val("");
            $("#ppatient_name").val("");
            $("#ppatient_lastname").val("");
            $("#ppatient_contact").val("");
            $("#pcodelast").val("");
            $("#pdate").val("");
            $("#pnumexp").val("");
            $("#pstudie").val("");
            $("#pservice").val("");
        }

        function agregar_tabla(){
            idservice=$("#pidservice").val();
            service=$("#pidservice option:selected").text();
            idstudy=$("#studies").val();
            study=$("#studies option:selected").text();
            comment=$("#pcomment").val();
            if (idservice !="" && idstudy !="" ){
                var fila='<tr class="selected" id="fila'+cont+'"><td><button type="button" class="btn btn-warning" onclick="eliminar('+cont+');">X</button></td><td><input type="hidden" name="idservice[]" value="'+idservice+'">'+service+'</td><td><input type="hidden" name="idstudy[]" value="'+idstudy+'">'+study+'</td><td><input type="hidden" name="comment[]" value="'+comment+'">'+comment+'</td></tr>';
                cont++;
                limpiar();
                evaluar();
                $('#detalles').append(fila);
            }else{
                alert("Error al ingresar el detalle de la cita, revise los datos de seleccionados.")
            }
        }

        function evaluar(){
            if (cont > 0){
                $("#btn_guardar").show();
            }else{
                $("#btn_guardar").hide();
            }
        }

        function eliminar(index){
            $("#fila" + index).remove();
            cont--;
            evaluar();
        }
        
    </script>
@endsection