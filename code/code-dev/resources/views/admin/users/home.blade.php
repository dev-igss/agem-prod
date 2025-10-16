@extends('admin.master')
@section('title','Agregar Usuario')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/users/1') }}" class="nav-link"><i class="fas fa-user-lock"></i> Usuarios</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-4">
                @if(kvfj(Auth::user()->permissions, 'user_add'))
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-plus-circle"></i> <strong>Agregar Usuario</strong></h2>
                        </div>

                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/usuario/agregar') }}">
                                @csrf

        
                                <label for="name"><strong><sup style="color: red;">(*)</sup>  IBM Empleado:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" name="ibm">
                                </div>

                                <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Nombre:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" name="name">
                                </div>

                                <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Apellidos:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span> 
                                    <input type="text" class="form-control" name="lastname">
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
                        <h2 class="title"><i class="fas fa-hospital-user"></i> <strong>Listado de Usuarios </strong></a>
                    </div>

                    <div class="inside">
                        <table id="table-modules" class="table table-striped table-hover mtop16">
                            <thead>
                                <tr>
                                    <td><strong> OPCIONES </strong></td>
                                    <td><strong> NOMBRE </strong></td>
                                    <td><strong> ROL </strong></td>
                                    <td><strong> ESTADO </strong></td>
                                </tr>
                            </thead>
                            <tbody>
                                @foreach($users as $user)
                                    <tr>
                                        <td>
                                            <div class="opts">
                                                @if(kvfj(Auth::user()->permissions, 'user_edit'))
                                                    <a href="{{ url('/admin/usuario/'.$user->id.'/editar') }}" data-toogle="tooltrip" data-placement="top" title="Ver"><i class="fas fa-eye"></i></a>
                                                @endif
                                                @if(kvfj(Auth::user()->permissions, 'user_permissions'))
                                                    <a href="{{ url('/admin/usuario/'.$user->id.'/permisos') }}" data-toogle="tooltrip" data-placement="top" title="Permisos"><i class="fas fa-cogs"></i></a>
                                                @endif
                                                @if(kvfj(Auth::user()->permissions, 'user_requests_out') && $user->role == 5)
                                                    <a href="{{ url('/admin/usuario/'.$user->id.'/solicitudes_fuera_de_tiempo') }}" data-toogle="tooltrip" data-placement="top" title="Solicitudes Fuera de Tiempo"><i class="fas fa-clock"></i></a>
                                                @endif
                                            </div>
                                        </td>
                                        <td>{{'IBM:'.$user->ibm.' - '.$user->name.' '.$user->lastname}}</td>
                                        <td>{{ getRoleUserArray(null, $user->role) }}</td>
                                        <td>{{ getUserStatusArray(null, $user->status) }}</td>
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
