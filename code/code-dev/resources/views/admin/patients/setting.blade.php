@extends('admin.master')
@section('title','Configuración de Citas')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/all') }}" class="nav-link"><i class="fas fa-columns"></i> Citas</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/add') }}" class="nav-link"><i class="fa fa-cogs"></i> Configuración de Citas</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        {!! Form::open(['url' => '/admin/paciente/configuracion']) !!}
        <div class="row">
        
            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Correlativo RX</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Iniciar:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    {!! Form::text('correlative_rx', $config->correlative_rx  ,['class'=>'form-control', 'Placeholder' => 'Ingrese una nueva cantida']) !!}
                                </div>
                            </div>
                        </div>                        
                    </div>
                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Correlativo USG</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Iniciar:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    {!! Form::text('correlative_usg', $config->correlative_usg  ,['class'=>'form-control', 'Placeholder' => 'Ingrese una nueva cantida']) !!}
                                </div>
                            </div>
                        </div>                  
                    </div>
                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Correlativo MMO</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Iniciar:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    {!! Form::text('correlative_mmo', $config->correlative_mmo ,['class'=>'form-control', 'Placeholder' => 'Ingrese una nueva cantida']) !!}
                                </div>
                            </div>
                        </div>                        
                    </div>
                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Correlativo DMO</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Iniciar:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    {!! Form::text('correlative_dmo', $config->correlative_dmo  ,['class'=>'form-control', 'Placeholder' => 'Ingrese una nueva cantida']) !!}
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
                        {!! Form::submit('Guardar', ['class'=>'btn btn-success', 'id'=>'btn_guardar']) !!}
                    </div>
                </div>                    
            </div>
        </div>

        {!! Form::close() !!}
    </div> 
@endsection