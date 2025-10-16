 @extends('patients_day.master')
@section('title','Citas Abiertas')
@section('content')
<section style="text-align: center;">
    <div class="home_action_bar shadow">
        <div class="row"> 
            <div class="col-md-3">
                <div class="categories" style="z-index: 999">
                    <a href="#"><i class="fas fa-stream"></i> Filtar</a>
                    <ul class="shadow">
                            <li>
                                <a href="{{ url('/citas_abiertas/todas') }}"> Todas las Citas </a>
                                <a href="{{ url('/citas_abiertas/rx') }}"> Citas de RX </a>
                                <a href="{{ url('/citas_abiertas/usg') }}"> Citas de USG </a>
                                <a href="{{ url('/citas_abiertas/mmo') }}"> Citas de MMO </a>
                                <a href="{{ url('/citas_abiertas/dmo') }}"> Citas de DMO </a>
                            </li>
                    </ul>
                </div>
            </div>

            <h4 style="margin-top: -35px; "><strong>Estudios Abiertos ({{ $total }})</strong></h4>
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
                             
                            <p class="card-text" style="margin-top:-15px;"><strong>Edad - Genero: </strong> {{ $a->patient->age.'a' }} -  @if($a->patient->gender == 0) Masculino @else Femenino @endif  </p>        
                            <p class="card-text" style="margin-top:-15px;"> <strong>Expediente: </strong> {{ $a->num_study }}</p>     
                            <p class="card-text">
                                <small> 
                                    <strong>Estudio(s): </strong> <br>
                                    @foreach($a->details as $det) 
                                        <span style="font-size: 0.90em;"><i class="fas fa-chevron-right"></i> {{ $det->study->name ?? 'Ninguno'  }} <br></span>
                                    @endforeach
                                    

                                </small>
                                <p class="card-text">
                                    <strong>Servicio: </strong> 
                                        {{ $a->service }}                                                
                                </p> 

                                <p class="card-text">
                                    <strong>Fecha: </strong> 
                                        {{ date('d-m-Y', strtotime($a->date)) }}                                                
                                </p> 
                                <p class="card-text">
                                    <strong>Hora de la Cita: </strong> 
                                    {{ \Carbon\Carbon::parse($a->schedule->hour_in)->format('H:i').' '.getHourType(null, $a->schedule->type) }}                                              
                                </p> 
                            </p>    
                                                       
                        </div>
                        <div class="card-footer text-muted">
                            <a href="{{ url('/citas_del_dia/'.$a->id.'/materiales') }}" data-action="materials" >Más Información</a>
                        </div>
                    </div>                    
                </div>
            @endforeach
        </div>    
    </div>
</div>


@endsection