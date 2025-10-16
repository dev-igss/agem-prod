@extends('admin.master')
@section('title','Agregar Usuario')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/users/1') }}" class="nav-link"><i class="fas fa-user-lock"></i> Usuarios</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/users/1') }}" class="nav-link"><i class="fas fa-user-lock"></i> Solicitudes Fuera de Tiempo: {{ $u->name.' '.$u->lastname}} (IBM: {{ $u->ibm}})</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-3">
                @if(kvfj(Auth::user()->permissions, 'user_requests_out'))
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-plus-circle"></i> <strong>Habilitar Nueva Solicitud</strong></h2>
                        </div>

                        <div class="inside">
                                <form method="POST" action="{{ url('/admin/usuario/'.$u->id.'/solicitudes_fuera_de_tiempo') }}">
                                @csrf
                            
                                <label for="name"><strong><sup style="color: red;">(*)</sup>  Jornada:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <select name="journey" class="form-select" aria-label="Default select example">
                                        @foreach($journeys as $j)
                                            <option value="{{$j->id}}">{{$j->name}}</option>
                                        @endforeach
                                    </select>
                                </div>

                                <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Cantidad de Dietas / Refacciones:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <input type="number" min="1" class="form-control" name="amount_diets">
                                </div>

                                <label for="name" class="mtop16"><strong><sup style="color: red;">(*)</sup> Tiempo a Asignar:</strong> </label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                    <select name="time" class="form-select" multiple aria-label="multiple select example">
                                        <option selected>Seleccione una opción</option>
                                        <option value="1">5 min.</option>
                                        <option value="2">10 min.</option>
                                        <option value="3">15 min.</option>
                                        <option value="4">20 min.</option>
                                        <option value="5">25 min.</option>
                                        <option value="6">30 min.</option>
                                    </select>
                                </div>

                                <button class="btn btn-success mt-4" type="submit">Guardar</button>
                            </form>

                        </div>
                    </div>
                @endif
            </div>

            <div class="col-md-9">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-hospital-user"></i> <strong>Listado de Solicitudes Habilitadas: </strong>{{ $u->name.' '.$u->lastname}} (IBM: {{ $u->ibm}}) </a>
                    </div>

                    <div class="inside">
                        <table id="table-modules" class="table table-striped table-hover mtop16" style="text-align:center;">
                            <thead>
                                <tr>
                                    <td><strong> FECHA HABILITADA</strong></td>
                                    <td><strong> JORNADA </strong></td>
                                    <td><strong> CANTIDAD DE DIETAS / REFACCIONES </strong></td>
                                    <td><strong> TIEMPO ASIGNADO</strong></td>
                                    <td><strong> ESTADO </strong></td>
                                </tr>
                            </thead>
                            <tbody>
                                @foreach($solicitudes as $s)
                                    <tr>
                                        <td>{{ $s->created_at->format('d/m/Y')}}</td>
                                        <td>{{ $s->journey->name }}</td>
                                        <td>{{ $s->amount_diets }}</td>
                                        <td>{{ getTimeRequestsOut(null, $s->time_available) }}</td>
                                        <td>{{ $s->status }}</td>
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
