@extends('admin.master')
@section('title','Horarios')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/horarios') }}" class="nav-link"><i class="fa fa-clock-o"></i> Horarios</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-4">
                @if(kvfj(Auth::user()->permissions, 'schedule_add'))
                    <div class="panel shadow">
                        <div class="header">
                            <h2 class="title"><i class="fas fa-plus-circle"></i><strong> Agregar Horario</strong></h2>                            
                        </div>

                        <div class="inside">
                            <form method="POST" action="{{ url('/admin/horario/agregar') }}">
                                @csrf
                                <div class="row">
                                    <div class="col-md-6">
                                        <label for="name"> <strong><sup style="color: red;">(*)</sup> Hora de Inicio: </strong></label>
                                        <div class="input-group">
                                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-keyboard"></i></span>
                                            <input type="datatime" class="form-control" value="{{\Carbon\Carbon::now()}}" name="hour_in">
                                        </div>
                                    </div>

                                    <div class="col-md-6">
                                        <label for="type"><strong><sup style="color: red;">(*)</sup> Hora de Finalización:</strong></label>
                                        <div class="input-group">
                                            <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>
                                            <input type="datatime" class="form-control" value="{{\Carbon\Carbon::now()}}" name="hour_out">
                                        </div>
                                    </div>
                                </div>
                                

                                

                                <label for="type"  class="mtop16"><strong><sup style="color: red;">(*)</sup> Jornada:</strong></label>
                                <div class="input-group">
                                    <span class="input-group-text" id="basic-addon1"><i class="fas fa-layer-group"></i></span>
                                    <select name="type" class="form-select" aria-label="Default select example">
                                    @foreach(getHourType('list', null) as $key => $value)
                                        
                                        <option value="{{ $key }}">{{ $value }}</option>
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
                        <h2 class="title"><i class="fas fa-heartbeat"></i><strong> Listado de Horarios</strong></h2>                       
                    </div>

                    <div class="inside">
                        <table id="table-modules" class="table table-bordered table-striped" style="background-color:#EDF4FB;">
                            <thead>
                                <tr>
                                    <td><strong>HORA INICIO</strong></td>
                                    <td><strong>HORA FINALIZACIÓN</strong></td>
                                </tr>
                            </thead>
                            <tbody>
                                @foreach($schedules as $s)
                                    <tr>
                                        
                                        <td>{{$s->hour_in.' '.getHourType(null, $s->type) }}</td>
                                        <td>{{$s->hour_out.' '.getHourType(null, $s->type)}}</td>
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