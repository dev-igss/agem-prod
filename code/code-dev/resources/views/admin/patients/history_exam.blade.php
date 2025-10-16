@extends('admin.master')
@section('title','Historial de Citas de Paciente')
<?php set_time_limit(0);
ini_set("memory_limit",-1);
ini_set('max_execution_time', 0); ?>

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/pacientes') }}" class="nav-link"><i class="fas fa-columns"></i> Pacientes</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/paciente/historial_citas') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Historial de Citas</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">

        <div class="row">
            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    

                    <div class="header">
                        <h2 class="title"><i class="fas fa-radiation-alt"></i><strong> Citas RX</strong></h2>
                    </div>

                    <div class="inside">   
                        @if(count($appointments_rx) > 0 )                                              
                            <ul>
                                @foreach($appointments_rx as $arx)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Fecha: </strong> {{ date('d-m-Y', strtotime($arx->date)) }} 
                                        @switch($arx->status)
                                            @case(0):
                                                &nbsp;<small><strong style="color: black; ">{{ getStatusAppointment(null, $arx->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(1):
                                                &nbsp;<small><strong style="color: orange; ">{{ getStatusAppointment(null, $arx->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(2):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $arx->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(3):
                                                &nbsp;<small><strong style="color: green; ">{{ getStatusAppointment(null, $arx->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(4):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $arx->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(5):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $arx->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(6):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $arx->status)  }}</strong></small>  
                                                <br>
                                            @break
                                        @endswitch
                                        
                                        @if($arx->date_old != NULL)
                                            <br>
                                            &nbsp;&nbsp;&nbsp;&nbsp;<small><strong> Fecha Inicial: </strong> {{ date('d-m-Y', strtotime($arx->date_old)) }} <br></small>
                                        @endif

                                        <strong> Estudio(s): </strong>
                                            <ul>
                                            @foreach($arx->details as $det)
                                                @if($det->study->is_doppler == 1)
                                                    <span style="font-size: 0.90em; color: red;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @else
                                                    <span style="font-size: 0.90em;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @endif
                                            @endforeach
                                            </ul>    
                                    </li>         
                                    <hr>                       
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1.5em; margin-left: 15px; font-weight: bold;" >¡No tiene citas en esta área!</label>
                        @endif
                    </div>

                </div>
            </div>

            <div class="col-md-3 d-flex"> 
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-radiation-alt"></i><strong> Citas USG</strong></h2>
                    </div>

                    <div class="inside"> 
                        @if(count($appointments_usg) > 0 )                                              
                            <ul>
                                @foreach($appointments_usg as $ausg)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Fecha: </strong> {{ date('d-m-Y', strtotime($ausg->date)) }} 
                                        @switch($ausg->status)
                                            @case(0):
                                                &nbsp;<small><strong style="color: black; ">{{ getStatusAppointment(null, $ausg->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(1):
                                                &nbsp;<small><strong style="color: orange; ">{{ getStatusAppointment(null, $ausg->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(2):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $ausg->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(3):
                                                &nbsp;<small><strong style="color: green; ">{{ getStatusAppointment(null, $ausg->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(4):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $ausg->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(5):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $ausg->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(6):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $ausg->status)  }}</strong></small>  
                                                <br>
                                            @break
                                        @endswitch
                                        
                                        @if($ausg->date_old != NULL)
                                            <br>
                                            &nbsp;&nbsp;&nbsp;&nbsp;<small><strong> Fecha Inicial: </strong> {{ date('d-m-Y', strtotime($ausg->date_old)) }} <br></small>
                                        @endif

                                        <strong> Estudio: </strong>    
                                        <ul>
                                            @foreach($ausg->details as $det)
                                                @if($det->study->is_doppler == 1)
                                                    <span style="font-size: 0.90em; color: red;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @else
                                                    <span style="font-size: 0.90em;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @endif
                                            @endforeach
                                        </ul>   
                                    </li>         
                                    <hr>                       
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1.5em; margin-left: 15px; font-weight: bold;" >¡No tiene citas en esta área!</label>
                        @endif
                        
                    </div>

                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    

                    <div class="header">
                        <h2 class="title"><i class="fas fa-radiation-alt"></i><strong> Citas MMO</strong></h2>
                    </div>

                    <div class="inside">                                                 
                        @if(count($appointments_mmo) > 0 )                                              
                            <ul>
                                @foreach($appointments_mmo as $ammo)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Fecha: </strong> {{ date('d-m-Y', strtotime($ammo->date)) }} 
                                        @switch($ammo->status)
                                            @case(0):
                                                &nbsp;<small><strong style="color: black; ">{{ getStatusAppointment(null, $ammo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(1):
                                                &nbsp;<small><strong style="color: orange; ">{{ getStatusAppointment(null, $ammo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(2):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $ammo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(3):
                                                &nbsp;<small><strong style="color: green; ">{{ getStatusAppointment(null, $ammo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(4):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $ammo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(5):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $ammo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(6):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $ammo->status)  }}</strong></small>  
                                                <br>
                                            @break
                                        @endswitch
                                        
                                        @if($ammo->date_old != NULL)
                                            <br>
                                            &nbsp;&nbsp;&nbsp;&nbsp;<small><strong> Fecha Inicial: </strong> {{ date('d-m-Y', strtotime($ammo->date_old)) }} <br></small>
                                        @endif

                                        <strong> Estudio: </strong>   
                                        <ul>
                                        @foreach($ammo->details as $det)
                                                @if($det->study->is_doppler == 1)
                                                    <span style="font-size: 0.90em; color: red;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @else
                                                    <span style="font-size: 0.90em;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @endif
                                            @endforeach
                                        </ul>    
                                    </li>         
                                    <hr>                       
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1.5em; margin-left: 15px; font-weight: bold;" >¡No tiene citas en esta área!</label>
                        @endif
                    </div>

                </div>
            </div>       
            
            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    

                    <div class="header">
                        <h2 class="title"><i class="fas fa-radiation-alt"></i><strong> Citas DMO</strong></h2>
                    </div>

                    <div class="inside">                                                 
                        @if(count($appointments_dmo) > 0 )                                              
                            <ul>
                                @foreach($appointments_dmo as $admo)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Fecha: </strong> {{ date('d-m-Y', strtotime($admo->date)) }} 
                                        @switch($admo->status)
                                            @case(0):
                                                &nbsp;<small><strong style="color: black; ">{{ getStatusAppointment(null, $admo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(1):
                                                &nbsp;<small><strong style="color: orange; ">{{ getStatusAppointment(null, $admo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(2):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $admo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(3):
                                                &nbsp;<small><strong style="color: green; ">{{ getStatusAppointment(null, $admo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(4):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $admo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(5):
                                                &nbsp;<small><strong style="color: yellow; ">{{ getStatusAppointment(null, $admo->status)  }}</strong></small>  
                                                <br>
                                            @break

                                            @case(6):
                                                &nbsp;<small><strong style="color: red; ">{{ getStatusAppointment(null, $admo->status)  }}</strong></small>  
                                                <br>
                                            @break
                                        @endswitch
                                        
                                        @if($admo->date_old != NULL)
                                            <br>
                                            &nbsp;&nbsp;&nbsp;&nbsp;<small><strong> Fecha Inicial: </strong> {{ date('d-m-Y', strtotime($admo->date_old)) }} <br></small>
                                        @endif
                                        <strong> Estudio: </strong>  
                                        <ul>
                                        @foreach($admo->details as $det)
                                                @if($det->study->is_doppler == 1)
                                                    <span style="font-size: 0.90em; color: red;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @else
                                                    <span style="font-size: 0.90em;"><i class="fa-solid fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                                @endif
                                            @endforeach
                                        </ul>     
                                    </li>         
                                    <hr>                       
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1.5em; margin-left: 15px; font-weight: bold;" >¡No tiene citas en esta área!</label>
                        @endif
                    </div>

                </div>
            </div>  

            
        </div>
    </div>
@endsection