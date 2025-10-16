@extends('admin.master')
@section('title','Agregar Paciente')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/pacientes') }}" class="nav-link"><i class="fas fa-columns"></i> Pacientes</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/paciente/historial_codigos_expedientes') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Historial de Citas</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">

        <div class="row">
            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    

                    <div class="header">
                        <h2 class="title"><i class="fas fa-barcode"></i><strong> Códigos RX</strong></h2>
                    </div>

                    <div class="inside">   
                        @if(count($codes_rx) > 0 )                                              
                            <ul>
                                @foreach($codes_rx as $crx)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Código: </strong> {{ $crx->code }} 
                                        @if($crx->status == '0')
                                            <small><strong style="color: green; "> Actual </strong></small>  
                                        @endif
                                        <br>
                                        <small><strong>Creado ó Asignado: </strong> {{ date('d-m-Y', strtotime($crx->created_at)) }}</small>
                                    </li>         
                                    <hr>                       
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1em; margin-left: 15px; font-weight: bold; text-align: center;" >¡No tiene código(s) generado(s) ó asignado(s)!</label>
                        @endif
                    </div>

                </div>
            </div>

            <div class="col-md-3 d-flex"> 
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-barcode"></i><strong> Códigos USG</strong></h2>
                    </div>

                    <div class="inside"> 
                        @if(count($codes_usg) > 0 )                                              
                            <ul>
                                @foreach($codes_usg as $cusg)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Código: </strong> {{ $cusg->code }} 
                                        @if($cusg->status == '0')
                                            <small><strong style="color: green; "> Actual </strong></small>     
                                        @endif
                                        <br>
                                        <small><strong>Creado ó Asignado: </strong> {{ date('d-m-Y', strtotime($cusg->created_at)) }}</small>
                                    </li>         
                                    <hr>                     
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1em; margin-left: 15px; font-weight: bold; text-align: center;" >¡No tiene código(s) generado(s) ó asignado(s)!</label>
                        @endif
                        
                    </div>

                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    

                    <div class="header">
                        <h2 class="title"><i class="fas fa-barcode"></i><strong> Códigos MMO</strong></h2>
                    </div>

                    <div class="inside">                                                 
                        @if(count($codes_mmo) > 0 )                                              
                            <ul>
                                @foreach($codes_mmo as $cmmo)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Código: </strong> {{ $cmmo->code }} 
                                        @if($cmmo->status == '0')
                                            <small><strong style="color: green; "> Actual </strong></small>      
                                        @endif
                                        <br>
                                        <small><strong>Creado ó Asignado: </strong> {{ date('d-m-Y', strtotime($cmmo->created_at)) }}</small>
                                    </li>         
                                    <hr>                      
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1em; margin-left: 15px; font-weight: bold; text-align: center;" >¡No tiene código(s) generado(s) ó asignado(s)!</label>
                        @endif
                    </div>

                </div>
            </div>       
            
            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    

                    <div class="header">
                        <h2 class="title"><i class="fas fa-barcode"></i><strong> Códigos DMO</strong></h2>
                    </div>

                    <div class="inside">                                                 
                        @if(count($codes_dmo) > 0 )                                              
                            <ul>
                                @foreach($codes_dmo as $cdmo)
                                    <li style="margin-left: 25px;"> 
                                        <strong> Código: </strong> {{ $cdmo->code }} 
                                        @if($cdmo->status == '0')
                                            <small><strong style="color: green; ">Actual</strong></small>    
                                        @endif
                                        <br>
                                        <small><strong>Creado ó Asignado: </strong> {{ date('d-m-Y', strtotime($cdmo->created_at)) }}</small>
                                    </li>         
                                    <hr>                      
                                @endforeach
                            </ul>
                        @else
                            <label  style="color: red; font-size: 1em; margin-left: 15px; font-weight: bold; text-align: center;" >¡No tiene código(s) generado(s) ó asignado(s)!</label>
                        @endif
                    </div>

                </div>
            </div>  

            
        </div>
    </div>
@endsection