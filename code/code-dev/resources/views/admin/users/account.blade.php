@extends('admin.master')
@section('title','Informacion de Cuenta')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/user/account/info') }}" class="nav-link"><i class="fas fa-id-card"></i> Información de Cuenta</a>
    </li>
@endsection

@section('content')
<div class="container-fluid">
    <div class="row">
        <div class="col-md-4">
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fas fa-fingerprint"></i> <strong>Cambiar Contraseña</strong></h2>
                </div>

                <div class="inside">
                    <form method="POST" action="{{ url('/admin/usuario/cuenta/cambiar/contrasena') }}">
                        @csrf
                        <div class="row">
                            <div class="col-md-12">
                                <label for="name"><strong>Contraseña Actual:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="password" class="form-control" name="apassword">
                                </div>
                            </div>
                        </div>

                        <div class="row mtop16">
                            <div class="col-md-12">
                                <label for="name"><strong>Nueva Contraseña:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="password" class="form-control" name="password">
                                </div>
                            </div>
                        </div>

                        <div class="row mtop16">
                            <div class="col-md-12">
                                <label for="name"><strong>Confirmar Contraseña:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="password" class="form-control" name="cpassword">
                                </div>
                            </div>
                        </div>

                        <div class="row mtop16">
                            <div class="col-md-12">
                                <button class="btn btn-primary mt-4" type="submit">Actualizar</button>
                            </div>
                        </div>
                    </form>
                </div>
            </div>
        </div>

        <div class="col-md-8">
            <div class="panel shadow">
                <div class="header">
                    <h2 class="title"><i class="fas fa-address-card"></i> <strong>Información del Usuario</strong></h2>
                </div>
                <div class="inside">
                    <form method="POST" action="{{ url('/cuenta/informacion/editar') }}">
                        @csrf
                        <div class="row">

                            <div class="col-md-4">
                                <label for="ibm"><strong>IBM:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{Auth::user()->ibm}}" name="ibm" disabled>
                                </div>
                            </div>

                            <div class="col-md-4">
                                <label for="name"><strong>Nombre:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{Auth::user()->name}}" name="name" disabled>
                                </div>
                            </div>

                            <div class="col-md-4">
                                <label for="lastname"><strong>Apellidos:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{Auth::user()->lastname}}" name="lastname" disabled>
                                </div>
                            </div>


                        </div>

                        <div class="row mtop16">
                            <div class="col-md-6">
                                <label for="email"><strong>Correo Institucional:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="email" class="form-control" value="{{Auth::user()->email}}" name="email" disabled>
                                </div>
                            </div>

                            <div class="col-md-6">
                                <label for="phone"><strong>Teléfono:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{Auth::user()->phone}}" name="phone" disabled>
                                </div>
                            </div>

                        </div>

                    </form>
                </div>
            </div>
        </div>
    </div>
</div>
@endsection
