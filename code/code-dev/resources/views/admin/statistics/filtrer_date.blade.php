@extends('admin.master')
@section('title','Panel Administrativo')

@section('content')
<div class="container-fluid mtop16">
        @if(kvfj(Auth::user()->permissions, 'dashboard_small_stats_month'))
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fa fa-filter"></i><strong> Filtrado Por Rango Fechas </strong></h2>
                </div>
            </div>

            <div class="row mtop16 ">
                <div class="col-md-12">
                    <div class="panel shadow">
                        <div class="inside">
                            {!! Form::open(['url'=>'/admin/reporte/filtrado/fechas']) !!}
                            <div class="row">

                                <div class="col-md-4">
                                    <label for="name"><strong>Desde:</strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        {!! Form::date('date_in', $date_in, ['class'=>'form-control']) !!}
                                    </div>
                                </div>

                                <div class="col-md-4">
                                    <label for="name"><strong>Hasta:</strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        {!! Form::date('date_out', $date_out, ['class'=>'form-control']) !!}
                                    </div>
                                </div>

                                
                                <div class="col-md-4" style="margin-top: 30px;">
                                    <div class="input-group">
                                        {!! Form::submit('Filtrar', ['class'=>'btn btn-warning']) !!}
                                        <a href="{{ url('/admin/reportes') }}" class="btn btn-info">Limpiar</a>
                                    </div>                  
                                </div>
                                
                            </div>
                            {!! Form::close() !!}
                        </div>
                    </div>
                </div>                               
            </div>           
        @endif
    </div>

    <div class="container-fluid mtop16">
        @if(kvfj(Auth::user()->permissions, 'dashboard_small_stats_month'))
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fas fa-chart-bar"></i><strong> Estadísticas Rápidas De Productividad Generales, Del {{\Carbon\Carbon::parse($date_in)->format('d-m-Y') }} Al {{\Carbon\Carbon::parse($date_out)->format('d-m-Y')}}</strong></h2>
                </div>
            </div>

            <div class="row mtop16">

                <div class="col-md-4">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-users"></i><strong> Citas Agendadas </strong></h2>
                        </div>
                        <div class="inside">
                            <div class="big_count">
                                {{ $citas_agen }}
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-md-4">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-users"></i><strong> Citas Atendidas </strong></h2>
                        </div>
                        <div class="inside">
                            <div class="big_count">
                                {{ $citas_aten }}
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-md-4">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-users"></i><strong> Citas Ausentes</strong></h2>
                        </div>
                        <div class="inside">
                            <div class="big_count">
                                0
                            </div>
                        </div>
                    </div>
                </div>

                
            </div>           
        @endif
    </div>

    <div class="container-fluid mtop16">
        @if(kvfj(Auth::user()->permissions, 'dashboard_prod_tecnicos_stats'))
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fas fa-chart-bar"></i><strong> Estadísticas De Productividad De los Tecnicos Por Area, Del {{\Carbon\Carbon::parse($date_in)->format('d-m-Y') }} Al {{\Carbon\Carbon::parse($date_out)->format('d-m-Y')}}</strong></h2>
                </div>
            </div>

            <div class="row mtop16 ">
                <div class="col-md-12">
                    <div class="panel shadow">
                        <div class="inside">
                            <table id="table-modules" class="table table-bordered table-striped" style=" text-align:center;">
                                <thead>
                                    <tr>
                                        <td ><strong>TÉCNICO</strong></td>
                                        <td ><strong>RX</strong></td>
                                        <td ><strong>USG</strong></td>
                                        <td ><strong>MMO</strong></td>
                                        <td ><strong>DMO</strong></td>
                                        <td ><strong>TOTAL</strong></td>

                                        
                                    </tr>
                                </thead>
                                <tbody>
                                    @foreach($tecnicos as $tec)                                
                                        <tr>                                
                                            <td>{{ $tec->ibm.' -  '.$tec->nombre }}</td>
                                            <td>
                                                @foreach($citas_tec_g_rx as $ctgrx)
                                                    @if($tec->id === $ctgrx->id_tecnico)
                                                        {{$ctgrx->total_citas}}
                                                    @endif
                                                @endforeach
                                            </td>
                                            <td>
                                                @foreach($citas_tec_g_usg as $ctgusg)
                                                    @if($tec->id === $ctgusg->id_tecnico)
                                                        {{$ctgusg->total_citas}}
                                                    @endif
                                                @endforeach                                          
                                            </td>
                                            <td>
                                                @foreach($citas_tec_g_mmo as $ctgmmo)
                                                    @if($tec->id === $ctgmmo->id_tecnico)
                                                        {{$ctgmmo->total_citas}}
                                                    @endif
                                                @endforeach                                          
                                            </td>
                                            <td>
                                                @foreach($citas_tec_g_dmo as $ctgdmo)
                                                    @if($tec->id === $ctgdmo->id_tecnico)
                                                        {{$ctgdmo->total_citas}}
                                                    @endif
                                                @endforeach
                                            </td>
                                            <td>       
                                                @foreach($citas_tec_g_total as $ctgtotal)
                                                    @if($tec->id === $ctgtotal->id_tecnico)
                                                        {{$ctgtotal->total_citas}}
                                                    @endif
                                                @endforeach                                    
                                            </td>
                                            
                                        </tr>
                                    @endforeach
                                                    
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>           
        @endif
    </div>

    




    

@endsection