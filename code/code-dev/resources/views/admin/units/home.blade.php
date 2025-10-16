@extends('admin.master')
@section('title','Unidades')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/unidades') }}" class="nav-link"><i class="fas fa-hospital-user"></i> Unidades</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-4">
                @if(kvfj(Auth::user()->permissions, 'unit_add'))
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-plus-circle"></i><strong> Agregar Unidad</strong></h2>
                        </div>

                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/unidad/agregar') }}">
                            @csrf
                                <label for="name"><strong><sup style="color: red;">(*)</sup> Nombre de Unidad:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" name="name">
                                </div>

                                <label for="name" class="mtop16"><strong> Codigo de Unidad:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="text" class="form-control" name="code">
                                </div>

                                <label class="mtop16"> <strong><sup style="color: red;">(*)</sup> Municipio de Ubicación </strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <select name="location_id" id="idsupplier" style="width: 90%;">
                                        @foreach ($locations as $l)                                    
                                            <option></option>
                                            <option value="{{$l->id}}">{{$l->name.' / '.$l->department}}</option>
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
                        <h2 class="title"><i class="fas fa-hospital-user"></i><strong> Unidades Hospitalarias </strong></a>
                    </div>

                    <div class="inside">
                        <table id="table-modules" class="table table-striped table-hover mtop16">
                            <thead>
                                <tr>
                                    <td width="140px"> <strong>OPCIONES</strong> </td>
                                    <td><strong> NOMBRE</strong> </td>
                                    <td><strong> UBICACION</strong> </td>
                                </tr>
                            </thead>
                            <tbody>
                                @foreach($units as $unit)
                                    <tr>
                                        <td>
                                        <div class="opts">
                                            @if(kvfj(Auth::user()->permissions, 'unit_edit'))
                                                <a href="{{ url('/admin/unidad/'.$unit->id.'/editar') }}" data-toogle="tooltrip" data-placement="top" title="Editar"><i class="fas fa-edit"></i></a>
                                            @endif
                                            @if(kvfj(Auth::user()->permissions, 'unit_delete'))
                                                <a href="#" data-action="borrar" data-path="admin/unidad" data-object="{{ $unit->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Borrar" ><i class="fas fa-trash"></i></a>    
                                            @endif
                                        </div>
                                        </td>
                                        <td>{{ $unit->name }}</td>
                                        <td>{{ $unit->location->name.' / '.$unit->location->department }}</td>
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
