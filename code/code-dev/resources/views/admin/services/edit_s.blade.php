@extends('admin.master')
@section('title','Servicios')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/environments/all') }}" class="nav-link"><i class="fa fa-object-group"></i> Servicios Generales</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/servicies_g/'.$environment->id.'/servicies') }}" class="nav-link"><i class="fa fa-object-group"></i> Servicios de: {{ $environment->name }}</a>
    </li>
@endsection

@section('content')
<div class="container-fluid">
    <div class="row">
        <div class="col-md-4">
            <div class="panel shadow">

                <div class="header">
                    <h2 class="title"><i class="fas fa-plus-circle"></i> Editar Servicio</h2>
                    
                </div>

                <div class="inside">
                    <form method="POST" action="{{ url('/admin/services_g/services/'.$services->id.'/edit') }}">
                    @csrf
                    
                    <input type="hidden" class="form-control" value="{{ $id }}" name="parent">

                        <label for="name"><strong><sup style="color: red;">(*)</sup> Nombre: </strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" value="{{ $services->name}}" name="name">
                        </div>

                        <button class="btn btn-success mt-4" type="submit">Guardar</button>
                    </form>
                </div>

            </div>
        </div>
    </div>
</div>

@endsection