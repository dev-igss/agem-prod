@extends('admin.master')
@section('title','Pacientes')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/pacientes') }}" class="nav-link"><i class="fas fa-users"></i> Pacientes</a>
    </li>
@endsection

@section('content')

<section>
    <div class="row">
        <div class="col-md-12">
            <div class="container-fluid">
                <div class="panel shadow">
                    <div class="header"> 
                        <h2 class="title"><i class="fas fa-users"></i> <strong>Filtro</strong></h2>
                        
                    </div>

                    <div class="inside" style="text-align: center; font-weight: bold;">
                        <div class="btn-group" role="group" aria-label="...">
                            <a href="{{url('/admin/pacientes/inicio')}}" type="button" class="btn btn-default">Inicio</a>
                            <a href="{{url('/admin/pacientes/a')}}" type="button" class="btn btn-default">A</a>
                            <a href="{{url('/admin/pacientes/b')}}" type="button" class="btn btn-default">B</a>
                            <a href="{{url('/admin/pacientes/c')}}" type="button" class="btn btn-default">C</a>
                            <a href="{{url('/admin/pacientes/ch')}}" type="button" class="btn btn-default">CH</a>
                            <a href="{{url('/admin/pacientes/d')}}" type="button" class="btn btn-default">D</a>
                            <a href="{{url('/admin/pacientes/e')}}" type="button" class="btn btn-default">E</a>
                            <a href="{{url('/admin/pacientes/f')}}" type="button" class="btn btn-default">F</a>
                            <a href="{{url('/admin/pacientes/g')}}" type="button" class="btn btn-default">G</a>
                            <a href="{{url('/admin/pacientes/h')}}" type="button" class="btn btn-default">H</a>
                            <a href="{{url('/admin/pacientes/i')}}" type="button" class="btn btn-default">I</a>
                            <a href="{{url('/admin/pacientes/j')}}" type="button" class="btn btn-default">J</a>
                            <a href="{{url('/admin/pacientes/k')}}" type="button" class="btn btn-default">K</a>
                            <a href="{{url('/admin/pacientes/l')}}" type="button" class="btn btn-default">L</a>
                            <a href="{{url('/admin/pacientes/m')}}" type="button" class="btn btn-default">M</a>
                            <a href="{{url('/admin/pacientes/n')}}" type="button" class="btn btn-default">N</a>
                            <a href="{{url('/admin/pacientes/ñ')}}" type="button" class="btn btn-default">Ñ</a>
                            <a href="{{url('/admin/pacientes/o')}}" type="button" class="btn btn-default">O</a>
                            <a href="{{url('/admin/pacientes/p')}}" type="button" class="btn btn-default">P</a>
                            <a href="{{url('/admin/pacientes/q')}}" type="button" class="btn btn-default">Q</a>
                            <a href="{{url('/admin/pacientes/r')}}" type="button" class="btn btn-default">R</a>
                            <a href="{{url('/admin/pacientes/s')}}" type="button" class="btn btn-default">S</a>
                            <a href="{{url('/admin/pacientes/t')}}" type="button" class="btn btn-default">T</a>
                            <a href="{{url('/admin/pacientes/u')}}" type="button" class="btn btn-default">U</a>
                            <a href="{{url('/admin/pacientes/v')}}" type="button" class="btn btn-default">V</a>
                            <a href="{{url('/admin/pacientes/w')}}" type="button" class="btn btn-default">W</a>
                            <a href="{{url('/admin/pacientes/x')}}" type="button" class="btn btn-default">X</a>
                            <a href="{{url('/admin/pacientes/y')}}" type="button" class="btn btn-default">Y</a>
                            <a href="{{url('/admin/pacientes/z')}}" type="button" class="btn btn-default">Z</a>
                            <a href="{{url('/admin/pacientes/borrados')}}" type="button" class="btn btn-default">Borrados</a>
                        </div>
                    </div>

                    <div class="col-md-9" style="margin-left: 200px; margin-bottom: 25px;">
                            <form method="POST" action="{{ url('/admin/paciente/busqueda') }}">
                                @csrf
                            <div class="input-group">
                                <select name="type_search" class="form-select col-2" aria-label="Default select example">
                                    <option selected>Elija una opción</option>
                                    <option value="1">Por Datos Paciente</option>
                                    <option value="2">Por Codigo Expediente</option>
                                </select>

                                <input type="text" class="form-control" name="search" placeholder="Realice una busqueda" required>
                                <button class="btn btn-primary mt-4" type="submit">Buscar</button>
                                <a href="{{ url('/admin/pacientes/inicio') }}" class="btn btn-warning">Limpiar</a>
                            </div>
                            </form>
                    </div>

                </div>
            </div>            
        </div>
    </div>
</section>

    <div class="container-fluid mtop16">
        <div class="panel shadow">

            <div class="header"> 
                <h2 class="title"><i class="fas fa-users"></i> <strong>Listado de Pacientes</strong></h2>
                <ul>
                    @if(kvfj(Auth::user()->permissions, 'patient_add'))
                        <li>
                            <a href="{{ url('/admin/paciente/agregar') }}" ><i class="fas fa-plus-circle"></i> Agregar Paciente</a>
                        </li>
                    @endif
                </ul>
            </div>

            <div class="inside">               

                <table id="table-modules" class="table table-striped table-hover mtop16">
                    <thead>
                        <tr>
                            <td><strong> OPCIONES </strong></td>
                            <td><strong> AFILIACION </strong></td>
                            <td><strong> APELLIDOS </strong></td>
                            <td><strong> NOMBRE </strong></td>                            
                            <td><strong> CONTACTO </strong></td>
                        </tr>
                    </thead>
                    <tbody>
                        @foreach($patients as $p)
                            <tr>
                                <td>
                                    <div class="opts">
                                        @if(kvfj(Auth::user()->permissions, 'patient_edit'))
                                            <a href="{{ url('/admin/paciente/'.$p->id.'/editar') }}" data-toogle="tooltrip" data-placement="top" title="Ver"><i class="fas fa-edit"></i></a>
                                        @endif
                                        @if(kvfj(Auth::user()->permissions, 'patient_history_exam'))
                                            <a href="{{ url('/admin/paciente/'.$p->id.'/historial_citas') }}" data-toogle="tooltrip" data-placement="top" title="Historial de Examenes"><i class="fas fa-book-medical"></i></a>
                                        @endif
                                        @if(kvfj(Auth::user()->permissions, 'patient_edit'))
                                            <a href="{{ url('/admin/paciente/'.$p->id.'/historial_codigos_expedientes') }}" data-toogle="tooltrip" data-placement="top" title="Historial de Codigos"><i class="fas fa-list"></i></a>
                                        @endif
                                        <!--@if($p->affiliation_principal != NULL)
                                            <a href="{{ url('/admin/paciente/'.$p->id.'/actualizar/afiliacion/principal/'.$p->affiliation_principal) }}" data-toogle="tooltrip" data-placement="top" title="Actualizar Afiliación Principal"><i class="fa fa-refresh"></i></a>
                                        @endif-->
                                        @if(kvfj(Auth::user()->permissions, 'patient_delete'))                                            
                                            <a href="#" data-action="borrar" data-path="admin/paciente" data-object="{{ $p->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Borrar" ><i class="fa-solid fa-trash"></i></a>    
                                        @endif
                                    </div>
                                </td>
                                <td>{{ getTypePatient(null, $p->type).': '.$p->affiliation}}</td>
                                <td>{{ $p->lastname }}</td>
                                <td>{{ $p->name }}</td>                                
                                <td>{{ $p->contact }}</td>
                            </tr>
                        @endforeach
                    </tbody>
                </table>
            </div>

        </div>
    </div>

    
@endsection