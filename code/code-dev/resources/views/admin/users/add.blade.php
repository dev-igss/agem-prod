@extends('admin.master')
@section('title','Agregar Usuario')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/users/1') }}" class="nav-link"><i class="fas fa-user-lock"></i> Usuarios</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/users/add') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Agregar Usuario</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="panel shadow">

            <div class="header">
                <h2 class="title"><i class="fas fa-plus-circle"></i> Agregar Usuario</h2>

            </div>

            <div class="inside">
                {!! Form::open(['url'=>'/admin/users/add']) !!}
                    <div class="row">

                        <div class="col-md-4">
                            <label for="name">Nombre:</label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                {!! Form::text('name', null, ['class'=>'form-control']) !!}
                            </div>
                            {!! $errors->first('name','<small style="color:red;">:message</small>') !!}
                        </div>

                        <div class="col-md-4">
                            <label for="lastname">Apellido:</label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                {!! Form::text('lastname', null, ['class'=>'form-control']) !!}
                            </div>
                            {!! $errors->first('lastname','<small style="color:red;">:message</small>') !!}
                        </div>

                        <div class="col-md-4">
                            <label for="ibm">IBM:</label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                {!! Form::text('ibm', null, ['class'=>'form-control']) !!}
                            </div>
                            {!! $errors->first('ibm','<small style="color:red;">:message</small>') !!}
                        </div>
                    </div>

                    <div class="row mtop16">
                        <div class="col-md-12">
                            {!! Form::submit('Guardar', ['class'=>'btn btn-success']) !!}
                        </div>
                    </div>

                {!! Form::close() !!}
            </div>

        </div>
    </div>
@endsection
