 @extends('patients_day.master')
@section('title','Citas del Día RX')
<?php set_time_limit(0);
ini_set("memory_limit",-1);
ini_set('max_execution_time', 0); ?>
@section('content')
<section style="text-align: center;">
    <div class="home_action_bar shadow">
        <div class="row"> 
            <h2><strong>Estudios Radiológicos Del Día: {{ $date }} </strong></h2>
            <br>

            @if($date == $today)
                <label class="mtop16" style="color: red; font-size: 2em; margin-left: 15px; font-weight: bold; " id="patient_msg" >¡Recuerde no dejar citas abiertas antes de terminar el día!</label>
            @endif
        </div>        
    </div>
</section>

<div class="container-fluid ">
    <div >       
        <div class="row d-flex">
            @foreach($appointments as $a)
                <div class="col-md-4 mtop16">                    
                    <div class="card text-center">
                        <div class="card-header" style="font-size: 1.2em; ">
                            <strong> Paciente: </strong>   <br>
                            {{ $a->patient->lastname.', '.$a->patient->name}} <br>
                        </div>

                        <div class="card-body" style="font-size: 1.2em;"> 
                             
                            <p class="card-text" style="margin-top:-15px;"><strong>Edad - Genero: </strong>  {{ $a->patient->age.'a' }} - @if($a->patient->gender == 0) Masculino @else Femenino @endif</p>        
                            <p class="card-text" style="margin-top:-15px;"> <strong>Expediente: </strong> {{ $a->num_study }}</p>     
                            <p class="card-text">
                                <small> 
                                    <strong>Estudio(s): </strong> 
                                    @foreach($a->details as $det)


                                        @if($det->study->is_doppler == 1)
                                            <span style="font-size: 0.90em; color: red;"><i class="fas fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                          @else
                                            <span style="font-size: 0.90em;"><i class="fas fa-chevron-right"></i> {{ $det->study->name }} <br></span>
                                        @endif

                                    @endforeach

                                </small>
                                <p class="card-text">
                                    <strong>Servicio: </strong> 
                                        {{ $a->service }}                                                
                                </p> 
                            </p>    
                                                       
                        </div>
                        <div class="card-footer text-muted">
                            <a href="{{ url('/citas_del_dia/'.$a->id.'/materiales') }}" data-action="materials" >Concluir Cita</a>
                        </div>
                    </div>                    
                </div>
            @endforeach
        </div>    
    </div>
</div>



@endsection