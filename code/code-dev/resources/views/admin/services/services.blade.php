@extends('admin.master')
@section('title','Servicios')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/servicios_general') }}" class="nav-link"><i class="fas fa-bed"></i> Servicios Generales</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/servicios_general/'.$environment->id.'/servicios') }}" class="nav-link"><i class="fa fa-object-group"></i> Servicios de: {{ $environment->name }}</a>
    </li>
@endsection

@section('content')
<div class="container-fluid">
    <div class="row">
        <div class="col-md-4">
        @if(kvfj(Auth::user()->permissions, 'service_add'))
            <div class="panel shadow">

                <div class="header">
                    <h2 class="title"><i class="fas fa-plus-circle"></i> Agregar Servicios</h2>
                    
                </div>

                <div class="inside">
                    <form method="POST" action="{{ url('/admin/servicios_general/servicios/agregar') }}">
                        @csrf
                        <input type="hidden" class="form-control" value="{{$id}}" name="parent_id">

                        <label for="name"><strong><sup style="color: red;">(*)</sup> Nombre: </strong></label>
                        <div class="input-group">
                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                            <input type="text" class="form-control" name="name">
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
                    <h2 class="title"><i class="fas fa-bed"></i> Servicios de: <b>{{ $environment->name.' - '.$environment->unit->name}}</b></h2>
                    
                </div>

                <div class="inside">
                    <table id="table-modules" class="table table-bordered table-striped" style="background-color:#EDF4FB;">
                        <thead>
                            <tr>
                                <td><strong>OPCIONES</strong></td>
                                <td><strong>NOMBRE</strong></td>
                            </tr>
                        </thead>
                        <tbody>
                            @foreach($services as $service)
                                <tr>
                                    <td>
                                        <div class="opts">
                                             
                                            @if(kvfj(Auth::user()->permissions, 'service_edit'))
                                                <a href="{{ url('/admin/servicios_general/servicios/'.$service->id.'/editar') }}" data-toogle="tooltrip" data-placement="top" title="Editar" class="edit"><i class="fas fa-edit"></i></a>
                                            @endif

                                        </div>
                                    </td>
                                    <td>{{ $service->name }}</td>
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