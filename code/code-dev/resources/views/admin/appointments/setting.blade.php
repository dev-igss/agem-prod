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
    <form method="POST" action="{{ url('/admin/cita/configuracion') }}">
    @csrf
        <div class="row">
        
            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Cantidad Citas RX</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Cantidad Citas A.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_rx_special_am}}" name="amount_rx_special_am" placeholder="Ingrese una nueva cantida">

                                </div>

                                <label for="name" class="mtop16"><strong>Cantidad Citas P.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_rx_special_pm}}" name="amount_rx_special_pm" placeholder="Ingrese una nueva cantida">
                                </div>
                            </div>
                        </div>                        
                    </div>
                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Cantidad Citas USG</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Cantidad Citas A.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_usg_am}}" name="amount_usg_am" placeholder="Ingrese una nueva cantida">
                                </div>

                                <label for="name" class="mtop16"><strong>Cantidad Citas P.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_usg_pm}}" name="amount_usg_pm" placeholder="Ingrese una nueva cantida">
                                </div>
                            </div>
                        </div>       
                        <div class="row">
                            <div class="col-md-12">
                                
                                <label for="name" class="mtop16"><strong>Cantidad Citas A.M (DOPPLER):</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_usg_doppler_am}}" name="amount_usg_doppler_am" placeholder="Ingrese una nueva cantida">
                                </div>

                                <label for="name" class="mtop16"><strong>Cantidad Citas P.M (DOPPLER):</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_usg_doppler_pm}}" name="amount_usg_doppler_pm" placeholder="Ingrese una nueva cantida">
                                </div>
                            </div>
                        </div>                  
                    </div>
                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Cantidad Citas MMO</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Cantidad Citas A.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_mmo_am}}" name="amount_mmo_am" placeholder="Ingrese una nueva cantida">
                                </div>

                                <label for="name" class="mtop16"><strong>Cantidad Citas P.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_mmo_pm}}" name="amount_mmo_pm" placeholder="Ingrese una nueva cantida">
                                </div>
                            </div>
                        </div>                        
                    </div>
                </div>
            </div>

            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Cantidad Citas DMO</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Cantidad Citas A.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_dmo_am}}" name="amount_dmo_am" placeholder="Ingrese una nueva cantida">
                                </div>

                                <label for="name" class="mtop16"><strong>Cantidad Citas P.M:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$config->amount_dmo_pm}}" name="amount_dmo_pm" placeholder="Ingrese una nueva cantida">
                                </div>
                            </div>
                        </div>                        
                    </div>
                </div>
            </div>

            
        </div>

        <div class="row mtop16">
        
            <div class="col-md-3 d-flex">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-list-ol"></i> <strong>Cantidad de Pacientes por Horario</strong></h2>
                    </div>

                    <div class="inside">                        
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Cantidad Pacientes:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{Config::get('agem.citas_configuradas')}}" name="citas_configuradas" placeholder="Ingrese una nueva cantida">
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
                        <button class="btn btn-success mt-4" id="btn_guardar" type="submit">Guardar</button>
                    </div>
                </div>                    
            </div>
        </div>

        </form>
    </div> 
@endsection