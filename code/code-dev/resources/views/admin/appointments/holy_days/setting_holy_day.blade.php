@extends('admin.master')
@section('title','Configuración de Dias Festivos')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/unidades') }}" class="nav-link"><i class="fas fa-hospital-user"></i> Días Festivos</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-4">
                @if(kvfj(Auth::user()->permissions, 'appointment_setting'))
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-plus-circle"></i><strong> Agregar Día Festivo</strong></h2>
                        </div>

                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/cita/configuracion/dias/festivos') }}">
                                @csrf
                                <label for="name"><strong><sup style="color: red;">(*)</sup> Nombre del Día Festivo:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" name="name" id="name" required>
                                </div>


                                
                                <label class="mtop16" for="name"><strong><sup style="color: red;">(*)</sup> Asigne la Fecha:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="date" class="form-control" name="holy_day" required>
                                </div>

                                <button class="btn btn-success mt-4" id="btn_guardar" type="submit">Guardar</button>
                            </form>
                        </div>
                    </div>
                @endif
            </div>

            <div class="col-md-8">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-hospital-user"></i><strong> Listado de Días Festivos </strong></a>
                    </div>

                    <div class="inside">
                        <table id="table-modules" class="table table-striped table-hover mtop16">
                            <thead>
                                <tr>
                                    <td width="140px"> <strong>OPCIONES</strong> </td>
                                    <td><strong> DÍA FESTIVO</strong> </td>
                                </tr>
                            </thead>
                            <tbody>
                                @foreach($config as $c)
                                    <tr>
                                        <td>
                                        <div class="opts">
                                            @if(kvfj(Auth::user()->permissions, 'appointment_setting'))
                                                <a href="{{ url('/admin/unidad/'.$c->id.'/editar') }}" data-toogle="tooltrip" data-placement="top" title="Editar"><i class="fas fa-edit"></i></a>
                                            @endif
                                            @if(kvfj(Auth::user()->permissions, 'appointment_setting'))

                                                <a href="#" data-action="borrar" data-path="admin/cita/configuracion/dias/festivos" data-object="{{ $c->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Borrar" ><i class="fas fa-trash"></i></a>    
                                            @endif
                                        </div>
                                        </td>
                                        <td>
                                            
                                            <span style="color: black; font-size: 18px;"> <b> {{ \Carbon\Carbon::parse($c->holy_day)->format('d-m-Y')  }}</b> </span> <br>
                                            <span style="color: black; font-size: 14px;"> <b>{{ $c->name}}</b> </span>
                                        </td>
                                    </tr>
                                @endforeach
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>

@endsection
