@extends('admin.master')
@section('title','Categorías')

@section('breadcrumb')
    <li class="breadcrumb-item">
    <a href="{{ url('/admin/units') }}" class="nav-link"><i class="fas fa-hospital-user"></i> Unidades</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">

        <div class="row">
            <div class="col-md-5">
                <div class="panel shadow">

                    <div class="header">
                        <h2 class="title"><i class="fas fa-edit"></i> Editar Unidad</h2>
                    </div>

                    <div class="inside">
                    <form method="POST" action="{{ url('/admin/unidad/'.$unit->id.'/editar') }}">
                    @csrf
                            
                            <label for="name"> <strong> Nombre de Unidad: </strong></label>
                            <div class="input-group">
                                <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                <input type="text" class="form-control" value="{{$unit->name}}" name="name">
                            </div>
                            
                            <label for="name" class="mtop16"><strong> Codigo de Unidad:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" value="{{$unit->code}}" name="code">
                                </div>
                            
                            <div class="row mtop16">
                                <div class="input-group">
                                    <button class="btn btn-success mt-4" type="submit">Guardar</button>&nbsp;
                                    <a href="{{ url('admin/units') }}" class="btn btn-danger mt-4">Cancelar</a>
                                </div>
                            </div>


                            </form>
                    </div>

                </div>
            </div>
    </div>

@endsection
