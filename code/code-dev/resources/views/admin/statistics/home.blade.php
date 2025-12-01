@extends('admin.master')
@section('title','Reportes')

@section('content')
    <div class="container-fluid mtop16">
        @if(kvfj(Auth::user()->permissions, 'report_month_areas'))
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fa fa-filter"></i><strong> Cuadres Mensuales </strong></h2>
                </div>
            </div>

            <div class="row mtop16 ">
                <div class="col-md-3">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-file-excel"></i><strong> Área de RX </strong></h2>
                        </div>

                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/reporte/estadisticas/rx') }}">
                            @csrf
                                <div class="inside">
                                    <label for="name"><strong>Seleccione el mes y año: </strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        <select name="month_rx"  class="form-select " aria-label="Default select example">
                                            @foreach(getMonths('list', null) as $key => $value)
                                                
                                                <option value="{{ $key }}">{{ $value }}</option>
                                            @endforeach
                                        </select>
                                        <input type="number" class="form-control" value="{{\Carbon\Carbon::now()->format('Y')}}" name="year_rx" min="2022" max="2099" step="1">

                                                                        
                                    </div>

                                </div>

                                <div class=" col-md-6 " >
                                    <div class="form-group">
                                        <input name="_token" value="{{ csrf_token() }}" type="hidden"></input>
                                        <button class="btn btn-primary" type="submit"> Generar</button>
                                    </div>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>    
                
                <div class="col-md-3">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-file-excel"></i><strong> Área de USG</strong></h2>
                        </div>

                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/reporte/estadisticas/usg') }}">
                            @csrf
                                <div class="inside">
                                    <label for="name"><strong>Seleccione el mes: </strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        <select name="month_usg" class="form-select " aria-label="Default select example">
                                            @foreach(getMonths('list', null) as $key => $value)
                                                
                                                <option value="{{ $key }}">{{ $value }}</option>
                                            @endforeach
                                        </select>
                                        <input type="number" class="form-control" value="{{\Carbon\Carbon::now()->format('Y')}}" name="year_usg" min="2022" max="2099" step="1">                        
                                    </div>

                                </div>

                                <div class=" col-md-6 " >
                                    <div class="form-group">
                                        <input name="_token" value="{{ csrf_token() }}" type="hidden"></input>
                                        <button class="btn btn-primary" type="submit"> Generar</button>
                                    </div>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>  

                <div class="col-md-3">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-file-excel"></i><strong> Área de MAMO </strong></h2>
                        </div>
                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/reporte/estadisticas/mamo') }}">
                            @csrf
                                <div class="inside">
                                    <label for="name"><strong>Seleccione el mes: </strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        <select name="month_mamo" class="form-select " aria-label="Default select example">
                                            @foreach(getMonths('list', null) as $key => $value)
                                                
                                                <option value="{{ $key }}">{{ $value }}</option>
                                            @endforeach
                                        </select>
                                        <input type="number" class="form-control" value="{{\Carbon\Carbon::now()->format('Y')}}" name="year_mamo" min="2022" max="2099" step="1">
                                        
                                    </div>

                                </div>

                                <div class=" col-md-6 " >
                                    <div class="form-group">
                                        <input name="_token" value="{{ csrf_token() }}" type="hidden"></input>
                                        <button class="btn btn-primary" type="submit"> Generar</button>
                                    </div>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>    
                
                <div class="col-md-3">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-file-excel"></i><strong> Área de DMO </strong></h2>
                        </div>
                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/reporte/estadisticas/dmo') }}">
                            @csrf
                           
                                <div class="inside">
                                    <label for="name"><strong>Seleccione el mes: </strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        <select name="month_dmo" class="form-select" aria-label="Default select example">
                                            @foreach(getMonths('list', null) as $key => $value)
                                                
                                                <option value="{{ $key }}">{{ $value }}</option>
                                            @endforeach
                                        </select>
                                        <input type="number" class="form-control" value="{{\Carbon\Carbon::now()->format('Y')}}" name="year_dmo" min="2022" max="2099" step="1">
                                                    
                                    </div>

                                </div>

                                <div class=" col-md-6 " >
                                    <div class="form-group">
                                        <input name="_token" value="{{ csrf_token() }}" type="hidden"></input>
                                        <button class="btn btn-primary" type="submit"> Generar</button>
                                    </div>
                                </div>
                            </form>
                        </div>
                    </div>
                </div> 
            </div>           
        @endif
    </div>

    <!--<div class="container-fluid mtop16">
        @if(kvfj(Auth::user()->permissions, 'report_tecnicos_individual'))
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fa fa-filter"></i><strong> Reporte de Tecnicos </strong></h2>
                </div>
            </div>

            <div class="row mtop16 ">
                <div class="col-md-12">
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-file-pdf"></i><strong> Generación de SPS-645 Por Tecnico </strong></h2>
                        </div>

                        <div class="inside">
                        <form method="POST" action="{{ url('/admin/reporte/tecnicos/individual') }}">
                        @csrf
                                <div class="inside">
                                    <label for="name"><strong>Tecnico &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    Mes
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                      Año 
                                      &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                                    &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Área
                                    </strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        <select name="idTecnico" id="idTecnico" class="form-select col-md-6 " aria-label="Default select example">
                                            <option selected>Seleccione una opción</option>
                                            @foreach($tecnicos as $tec)
                                                <option value="{{$tec->id}}" >{{'IBM: '.$tec->ibm.' - '.$tec->nombre}}</option>
                                            @endforeach
                                        </select>
                                        <select name="month_r_tec" class="form-select col-md-2" aria-label="Default select example">
                                            @foreach(getMonths('list', null) as $key => $value)
                                                
                                                <option value="{{ $key }}">{{ $value }}</option>
                                            @endforeach
                                        </select>
                                        <input type="number" class="form-control col-md-2" value="{{\Carbon\Carbon::now()->format('Y')}}" name="year_r_tec" min="2022" max="2099" step="1">   
                                        <select name="area" class="form-select col-md-2" aria-label="Default select example">
                                                                                         
                                            <option value="0" selected>RX</option>
                                            <option value="3">MMO</option>
                                            <option value="4">DMO</option>
                                            <option value="5">Todas</option>
                                        </select>                              
                                    </div>

                                </div>

                                <div class=" col-md-6 " >
                                    <div class="form-group">
                                        <input name="_token" value="{{ csrf_token() }}" type="hidden"></input>
                                        <button class="btn btn-primary" type="submit" target="_blank"> Generar</button>
                                    </div>
                                </div>
                                </form>
                        </div>
                    </div>
                </div>   
            </div>           
        @endif
    </div>

    -->

    @if(kvfj(Auth::user()->permissions, 'report_dates'))
        <div class="container-fluid mtop16">
            
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fa fa-filter"></i><strong> Filtrado Por Rango Fechas </strong></h2>
                </div>
            </div>

            <div class="row mtop16 ">
                <div class="col-md-12">
                    <div class="panel shadow">
                        <div class="inside">
                        <form method="POST" action="{{ url('/admin/reporte/filtrado/fechas') }}">
                        @csrf
                            <div class="row">

                                <div class="col-md-4">
                                    <label for="name"><strong>Desde:</strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        <input type="date" class="form-control" name="date_in">
                                    </div>
                                </div>

                                <div class="col-md-4">
                                    <label for="name"><strong>Hasta:</strong></label>
                                    <div class="input-group">
                                        <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                        <input type="date" class="form-control" name="date_out">
                                    </div>
                                </div>

                                
                                <div class="col-md-4" style="margin-top: 30px;">
                                    <div class="input-group">
                                        <button class="btn btn-warning mt-4" type="submit">Filtrar</button>
                                    </div>                  
                                </div>
                                
                            </div>
                            </form>
                        </div>
                    </div>
                </div>                               
            </div> 
        </div>

        <div class="container-fluid mtop16">
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fas fa-chart-bar"></i><strong> Estadísticas Rápidas De Productividad Generales </strong></h2>
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
                                0
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
                                0
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
        </div>

        <div class="container-fluid mtop16">
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fas fa-chart-bar"></i><strong> Estadísticas De Productividad De los Tecnicos Por Área </strong></h2>
                </div>
            </div>

            <div class="row mtop16 ">
                <div class="col-md-12">
                    <div class="panel shadow">
                        <div class="inside">
                            <table id="table-modules2" class="table table-bordered table-striped" style=" text-align:center;">
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
                                    
                                                    
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>                     
        </div>
    @endif

    




    

@endsection