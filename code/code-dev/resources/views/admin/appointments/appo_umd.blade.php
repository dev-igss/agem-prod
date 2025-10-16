
@extends('admin.master')
@section('title','Citas')
<?php set_time_limit(0);
ini_set("memory_limit",-1);
ini_set('max_execution_time', 0); ?>

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/citas/umd') }}" class="nav-link"><i class="fas fa-calendar-alt"></i> Citas U.M.D</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="panel shadow">

            <div class="header">
                <h2 class="title"><i class="fas fa-calendar-alt"></i><strong> Listado de Citas Ultrasonido - Mamografia - Densitometria</strong> </h2>
                <ul>
                    
                    <ul>
                        <li>
                            <div class="btn-group">
                                <button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown">
                                        <i class="fas fa-filter"></i>  Filtrado Estados <span class="caret"></span>
                                </button>

                                <ul class="dropdown-menu" role="menu">
                                    <li><a href="{{ url('/admin/cita/calendario/rx') }}">Agendada</a></li>
                                    <li><a href="{{ url('/admin/cita/calendario/umd') }}"> En Atención</a></li>
                                    <li><a href="{{ url('/admin/cita/calendario') }}"> Reprogramada</a></li>
                                    <li><a href="{{ url('/admin/cita/calendario') }}"> Ausente</a></li>
                                    <li><a href="{{ url('/admin/cita/calendario') }}"> Finalizada</a></li>
                                </ul>
                            </div>
                        </li>
                    </ul>
                    @if(kvfj(Auth::user()->permissions, 'appointment_calendar')) 
                        <ul>
                            <li>
                                <div class="btn-group">
                                    <button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown">
                                            <i class="fa fa-calendar"></i>  Ver Calendario <span class="caret"></span>
                                    </button>

                                    <ul class="dropdown-menu" role="menu">
                                        <li><a href="{{ url('/admin/cita/calendario/umd') }}"> USG-MMO-DMO</a></li>
                                    </ul>
                                </div>
                            </li>
                        </ul>
                    @endif
                    @if(kvfj(Auth::user()->permissions, 'appointment_setting'))
                        <li>
                            <a href="{{ url('/admin/cita/configuracion/dias/festivos') }}" ><i class="fas fa-calendar-day"></i> Días Festivos</a>
                        </li>
                    @endif
                    @if(kvfj(Auth::user()->permissions, 'appointment_setting'))
                        <li>
                            <a href="{{ url('/admin/cita/configuracion') }}" ><i class="fa fa-cogs"></i> Configuración de Citas</a>
                        </li>
                    @endif
                    @if(kvfj(Auth::user()->permissions, 'appointment_add'))
                        <li>
                            <a href="{{ url('/admin/cita/agregar') }}" ><i class="fas fa-plus-circle"></i> Agendar Cita</a>
                        </li>
                    @endif
                    
                </ul>
            </div>

            <div class="inside">  
                    <form method="POST" action="{{ url('/admin/cita/busqueda') }}">
                    @csrf
                        <div class="row" >
                            <div class="col-md-5" >                                
                                <div class="input-group">
                                <input type="date" class="form-control" name="search_date" placeholder="Realice una busqueda fecha">
                                </div>                                
                            </div>

                            <div class="col-md-5" >                                
                                <div class="input-group">
                                <select name="type_patient" id="patient_type" class="form-select col-md-2" aria-label="Default select example">
                                    @foreach(getTypePatient('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
                                    @endforeach
                                </select>
                                    <input type="text" class="form-control" name="search_patient" placeholder="Realice una busqueda por paciente">   
                                </div>     
                                
                                <input type="hidden" class="form-control" name="type_appo" value="2" > 
                            </div>

                            <div class="col-md-2" style="margin-right: -100px;">                                
                                <div class="input-group">
                                    <button class="btn btn-primary" type="submit">Buscar</button>                                  
                                </div>                                
                            </div>
                        </div>
                        
                    </form> 

                    <br>

                    <span><strong>Conteo de Citas Registradas del Día (Todos los servicios): </strong></span>
                    <br>

                    <span><strong><i class="fas fa-chevron-right"></i> Citas de USG A.M: </strong> {{ $app_usg_am }}</span> &nbsp&nbsp&nbsp&nbsp
                    <span><strong><i class="fas fa-chevron-right"></i> Citas de USG P.M: </strong> {{ $app_usg_pm }}</span> &nbsp&nbsp&nbsp&nbsp
                    <span><strong><i class="fas fa-chevron-right"></i> Citas de MMO A.M: </strong> {{ $app_mmo_am }}</span> &nbsp&nbsp&nbsp&nbsp
                    <span><strong><i class="fas fa-chevron-right"></i> Citas de MMO P.M: </strong> {{ $app_mmo_pm }}</span> &nbsp&nbsp&nbsp&nbsp
                    <span><strong><i class="fas fa-chevron-right"></i> Citas de DMO A.M: </strong> {{ $app_dmo_am }}</span> &nbsp&nbsp&nbsp&nbsp
                    <span><strong><i class="fas fa-chevron-right"></i> Citas de DMO P.M: </strong> {{ $app_dmo_pm }}</span>
                
                        
                <hr>      

                <table id="table-modules" class="table table-striped table-hover mtop16">
                    <thead>
                        <tr>
                            <td><strong> OPCIONES </strong></td>
                            <td><strong> FECHA </strong></td>
                            <td><strong> PACIENTE </strong></td>
                            <td><strong> SERVICIO </strong></td>
                            <td><strong> ESTUDIOS A REALIZAR</strong></td>
                            <td><strong> ESTADO </strong></td>
                        </tr> 
                    </thead>
                    <tbody>
                        @foreach($appointments as $a)
                            <tr>
                                <td>
                                    <div class="opts">
                                        @if(kvfj(Auth::user()->permissions, 'appointment_reschedule'))
                                            @if($a->status == '0')
                                                <a href="#" data-action="reprogramar" data-path="admin/cita" data-object="{{ $a->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Reprogramación" ><i class="fas fa-calendar-alt"></i></a>
                                            @endif
                                            @if($a->status == '5')
                                                <a href="#" data-action="reprogramacion_forzada" data-path="admin/cita" data-object="{{ $a->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Reprogramación Forzada" ><i class="fas fa-calendar-alt"></i></a>
                                            @endif
                                            @if($a->status == '0' || $a->status == '4' || $a->status == '5')
                                                <a href="#" data-date="{{$a->date}}" data-exam="{{$a->area}}" data-action="cambio_horario" data-path="admin/cita" data-object="{{ $a->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Cambio de Horario" ><i class="fas fa-clock"></i></a>
                                            @endif
                                        @endif

                                        <!--@if(kvfj(Auth::user()->permissions, 'appointment_reschedule'))
                                            @if($a->status == '0')
                                            <a href="{{ url('/admin/cita/'.$a->id.'/materiales') }}"  title="Constancia de Cita" ><i class="fa fa-file-text"></i></a>
                                            @endif
                                        @endif-->

                                        @if(kvfj(Auth::user()->permissions, 'appointment_patients_status'))
                                            @if($a->status == '0' || $a->status == '4')
                                                <a href="#" data-action="paciente_presente" data-path="admin/cita" data-object="{{ $a->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Paciente Presente"><i class="fas fa-calendar-check"></i></a>
                                                <a href="#" data-action="paciente_ausente" data-path="admin/cita" data-object="{{ $a->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Paciente Ausente"><i class="fas fa-calendar-times"></i></a>
                                            @endif
                                        @endif

                                        @if(kvfj(Auth::user()->permissions, 'appointment_materials'))
                                            @if($a->status == '3')
                                                <a href="{{ url('/admin/cita/'.$a->id.'/materiales') }}"  title="Materiales Usados" ><i class="fas fa-x-ray"></i></a>
                                            @endif
                                        @endif

                                        @if(kvfj(Auth::user()->permissions, 'appointment_materials'))
                                            @if($a->status == '3')
                                                <a href="{{ url('/admin/cita/'.$a->id.'/informe_al_patrono') }}" target="_blank" title="Informe al Patrono" ><i class="fa fa-file"></i></a>
                                            @endif
                                        @endif

                                        @if(kvfj(Auth::user()->permissions, 'appointment_delete'))
                                            @if($a->status == '0')
                                                <a href="#" data-action="borrar" data-path="admin/cita" data-object="{{ $a->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Borrar" ><i class="fas fa-trash"></i></a>    
                                            @endif
                                        @endif
                                    </div>
                                </td>
                                <td>
                                    {{ \Carbon\Carbon::parse($a->date)->format('d-m-Y') }} <br>
                                    @if($a->schedule_id != NULL)
                                        <small><strong>Horario: </strong> {{ \Carbon\Carbon::parse($a->schedule->hour_in)->format('H:i').' '.getHourType(null, $a->schedule->type) }}</small><br>
                                    @else
                                        <small><strong>Horario: </strong> Sin Asignar</small><br>
                                    @endif
                                    <small> @if($a->area) {{ getExamB(null, $a->area) }} @else nulo!! @endif - {{ getTypeAppointment(null, $a->type)  }} </small>
                                </td>
                                <td> 
                                    {{ $a->patient->name.' '.$a->patient->lastname }} <br>
                                    <span>AF. {{ $a->patient->affiliation }}</span> <br>                                    
                                    <small>Expediente. {{ $a->num_study }}</small>
                                </td>
                                <td>{{ $a->service }} </td>
                                <td> 
                                    
                                    @foreach($a->details as $det)


                                        @if($det->study->is_doppler == 1)
                                            <span style="font-size: 0.90em; color: red;"><i class="fas fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                        @else
                                            <span style="font-size: 0.90em;"><i class="fas fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                        @endif

                                    @endforeach
                                </td>
                                <td>
                                    @switch($a->status)
                                        @case(0)
                                            <small><strong style="color:black; ">{{ getStatusAppointment(null, $a->status)  }}</strong></small>  
                                        @break

                                        @case(1)
                                            <small><strong style="color:orange; ">{{ getStatusAppointment(null, $a->status)  }}</strong></small>  
                                        @break

                                        @case(2)
                                            <small><strong style="color:red; ">{{ getStatusAppointment(null, $a->status)  }}</strong></small>  
                                        @break

                                        @case(3)
                                            <small><strong style="color:green; ">{{ getStatusAppointment(null, $a->status)  }}</strong></small>  
                                        @break

                                        @case(4)
                                            <small><strong style="color:yellow; ">{{ getStatusAppointment(null, $a->status)  }}</strong></small>  
                                        @break

                                        @case(5)
                                            <small><strong style="color:yellow; ">{{ getStatusAppointment(null, $a->status)  }}</strong></small>  
                                        @break

                                        @case(6)
                                            <small><strong style="color:red; ">{{ getStatusAppointment(null, $a->status)  }}</strong></small>  
                                        @break
                                    @endswitch                                    
                                    @if($a->status == '3')
                                        <p>
                                            <small style=" font-size: 0.95em;">
                                                @if($a->ibm_tecnico_2 == NULL)
                                                    <strong >Tecnico: </strong> {{ $a->tecnico1->name.' '.$a->tecnico1->lastname }}
                                                @else
                                                    <strong >Tecnicos: </strong> {{ $a->tecnico1->name.' '.$a->tecnico1->lastname }} <br>
                                                    {{ $a->tecnico2->name.' '.$a->tecnico2->lastname }}
                                                @endif
                                            </small>
                                        </p>
                                    @endif
                                </td>
                            </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>

        </div>
    </div>


@endsection