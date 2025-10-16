@extends('admin.master')
@section('title','Servicios')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/servicios_general') }}" class="nav-link"><i class="fas fa-bed"></i> Servicios General</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-4">
                @if(kvfj(Auth::user()->permissions, 'serviceg_add'))
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-plus-circle"></i><strong> Agregar Servicio General</strong></h2>
                        </div>

                        <div class="inside">
                        <form method="POST" action="{{ url('/admin/servicios_general/agregar') }}">
                        @csrf
                                <label for="name"> <strong><sup style="color: red;">(*)</sup> Nombre: </strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" name="name">
                                </div>

                                <label for="unit_id"  class="mtop16"><strong><sup style="color: red;">(*)</sup> Unidad Hospitalaria:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>
                                    <select name="unit_id" class="form-select" aria-label="Default select example">
                                        <option value="">Seleccione una opción</option>
                                        @foreach($units as $u)
                                            <option value="{{$u->id}}">{{$u->name}}</option>
                                        @endforeach
                                    </select>
                                </div>

                                <button class="btn btn-success mt-4" type="submit">Guardar</button>
                            </form>
                        </div>
                    </div>
                @endif
            </div>

            <div class="col-md-8">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-bed"></i><strong> Servicios Generales</strong></h2>
                        
                    </div>

                    <div class="inside">
                        <table id="table-modules" class="table table-bordered table-striped" style="background-color:#EDF4FB;">
                            <thead>
                                <tr>
                                    <td><strong>OPCIONES</strong></td>
                                    <td><strong>NOMBRE</strong></td>
                                    <td><strong>UNIDAD HOSPITALARIA</strong></td>
                                </tr>
                            </thead>
                            <tbody>
                                @foreach($environment as $env)
                                    <tr>
                                        <td>
                                            <div class="opts">

                                                @if(kvfj(Auth::user()->permissions, 'serviceg_edit'))
                                                    <a href="{{ url('/admin/servicios_general/'.$env->id.'/editar') }}"  title="Editar"><i class="fas fa-edit"></i></a>
                                                @endif                                             
                                                
                                                @if(kvfj(Auth::user()->permissions, 'service_list'))
                                                    <a href="{{ url('/admin/servicios_general/'.$env->id.'/servicios') }}"  title="Servicios"><i class="fas fa-list-ul"></i></a>
                                                @endif  

                                            </div>
                                        </td>
                                        <td>{{$env->name}}</td>
                                        <td>{{$env->unit->name}}</td>
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
