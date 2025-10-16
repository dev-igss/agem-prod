@extends('admin.master')
@section('title','Editar Usuario')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/users/all') }}" class="nav-link"><i class="fas fa-user-lock"></i> Usuarios</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="page_user">
            <div class="row">
                <div class="col-md-4 d-flex">
                    <div class="panel shadow">
                        <div class="header"> 
                            <h2 class="title"><i class="fas fa-info-circle"></i><strong> Información Actual</strong></h2>
                        </div>

                        <div class="inside">
                            <div class="mini_profile">
                                <div class="info">
                                    <span class="title"><i class="fas fa-user-circle"></i> Nombre:</span>
                                    <span class="text">{{ $u->name.' '.$u->lastname}}</span>

                                    <span class="title"><i class="fas fa-id-card"></i> IBM:</span>
                                    <span class="text">{{ $u->ibm}}</span>

                                    <span class="title"><i class="fas fa-user-tie"></i> Estado del Usuario:</span>
                                    <span class="text">{{ getUserStatusArray(null, $u->status) }}</span>

                                    <span class="title"><i class="far fa-calendar-alt"></i> Fecha de Registró:</span>
                                    <span class="text">{{ $u->created_at }}</span>

                                    <span class="title"><i class="fas fa-user-shield"></i> Rol de Usuario:</span>
                                    <span class="text">{{ getRoleUserArray(null, $u->role) }}</span>
                                </div>

                                @if(kvfj(Auth::user()->permissions, 'user_banned'))
                                    @if($u->status == '0')
                                        <a href="{{ url('/admin/usuario/'.$u->id.'/suspender') }}" class="btn btn-success">Activar Usuario</a>
                                    @else
                                        <a href="{{ url('/admin/usuario/'.$u->id.'/suspender') }}" class="btn btn-warning">Suspender Usuario</a>
                                    @endif
                                @endif
                            </div>
                        </div>
                    </div>
                </div>

                <div class="col-md-4 d-flex">

                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-edit"></i><strong> Editar Información</strong></h2>
                        </div>

                        <div class="inside">

                            @if(kvfj(Auth::user()->permissions, 'user_edit'))
                                <form method="POST" action="{{ url('/admin/usuario/'.$u->id.'/editar') }}">
                                    @csrf
                                    <div class="row">
                                        <div class="col-md-12">
                                            <label for="name" ><strong>Nombre:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                                <input type="text" class="form-control" name="name">
                                            </div>

                                            <label for="lastname" class="mtop16"><strong>Apellidos:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                                <input type="text" class="form-control" name="lastname">
                                            </div>

                                            <label for="module" class="mtop16"><strong>Tipo de Usuario:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>
                                                <select class="form-select" aria-label="Default select example">
                                                        <option value="1">Administrador de Unidad</option>
                                                        <option value="2">Jefe de Alimentacion</option>
                                                        <option value="3">Encargado de Dietistas</option>
                                                        <option value="4">Dietista</option>
                                                        <option value="5">Jefe de Servicio</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>

                                    <div class="row mtop16">
                                        <div class="col-md-12">
                                        <button class="btn btn-success mt-4" type="submit">Guardar</button>
                                        </div>
                                    </div>
                                </form>
                            @endif
                        </div>
                    </div>
                </div>

                @if(kvfj(Auth::user()->permissions, 'user_reset_password'))
                    <div class="col-md-4 d-flex">
                        <div class="panel shadow mtop32">
                            <div class="header">
                                <h2 class="title"><i class="fas fa-fingerprint"></i><strong> Restablecer Contraseña</strong></h2>
                            </div>
                            <div class="inside">
                                <form method="POST" action="{{ url('/admin/usuario/'.$u->id.'/reiniciar_contrasena') }}">
                                @csrf
                                    <div class="row">
                                        <div class="col-md-12">
                                            <label for="name"><strong>Nueva Contraseña:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                                <input type="password" class="form-control" name="password">
                                            </div>
                                            <!--{!! $errors->first('password','<small style="color:red;">:message</small>') !!}-->
                                        </div>
                                    </div>

                                    <div class="row mtop16">
                                        <div class="col-md-12">
                                            <label for="name"><strong>Confirmar Contraseña:</strong></label>
                                            <div class="input-group">
                                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                                <input type="password" class="form-control" name="cpassword">
                                            </div>
                                            <!--{!! $errors->first('cpassword','<small style="color:red;">:message</small>') !!}-->
                                        </div>
                                    </div>

                                    <div class="row mtop16">
                                        <div class="col-md-12">
                                        <button class="btn btn-primary mt-4" type="submit">Restablecer</button>
                                        </div>
                                    </div>
                                </form>
                            </div>
                        </div>
                    </div>
                @endif
            </div>
        </div>
    </div>
@endsection
