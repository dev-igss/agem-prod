@extends('admin.master')
@section('title','Materiales Usados')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/all') }}" class="nav-link"><i class="fas fa-columns"></i> Citas</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/add') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Material Usados</a>
    </li>
@endsection

@section('content')
<div class="container-fluid">
    <div class="row">
        <div class="col-md-3">
            <div class="panel shadow">
                <div class="header"> 
                    <h2 class="title"><i class="fas fa-clipboard-list"></i><strong> Listado de Estudios de la Cita</strong>  </h2>
                </div>

                <div class="inside">
                    <ul>
                        @foreach($detalles as $det)
                            <li style="margin-left: 25px;"> 
                                <strong> Estudio: </strong> {{ $det->study->name }}   <br>
                                <a href="#" data-action="registro_materiales" data-path="admin/cita" data-object="{{ $det->idappointment }}" data-study="{{ $det->idstudy }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" >Asignar Material</a>
                            </li>           
                            <hr>                       
                        @endforeach                        
                    </ul>
                </div>
            </div>
        </div>

        <div class="col-md-9">
            <div class="panel shadow">
                <div class="header"> 
                    <h2 class="title"><i class="fas fa-clipboard-list"></i><strong> Listado de Materiales Usados en Cita</strong>  </h2>
                </div>

                <div class="inside">
                    <table id="table-modules" class="table table-striped table-hover">
                        <thead>
                            <tr>
                                <td><strong> OPCIONES </strong></td>
                                <td><strong> FECHA DE REGISTRO </strong></td>
                                <td><strong> ESTUDIO</strong></td>
                                <td><strong> MATERIAL </strong></td>
                                <td><strong> CANTIDAD</strong></td>
                            </tr>
                        </thead>
                        <tbody>
                            @foreach($materials as $ma)
                                <tr>
                                    <td>
                                        <div class="opts">   
                                            @if($ma->status != '1')
                                                <a href="#" data-action="material_desechado" data-path="admin/cita/materiales" data-object="{{ $ma->id }}" class="btn-deleted" data-toogle="tooltrip" data-placement="top" title="Material Descartado" ><i class="fas fa-trash"></i></a>                                                                                       
                                            @endif
                                        </div>
                                    </td>
                                    <td>{{ $ma->created_at }}</td>
                                    <td>{{ $ma->study->name }}</td>
                                    <td>
                                        {{ getMaterials(null, $ma->material)  }} 
                                        @if($ma->status == '1')
                                            <br>
                                            <span style="color: red; font-weight: bold;" ><small>Desechado</small></span>
                                        @endif
                                    </td>
                                    <td>{{ $ma->amount }}</td>
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