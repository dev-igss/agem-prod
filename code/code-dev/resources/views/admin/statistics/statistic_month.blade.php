@extends('admin.master')
@section('title','Bitacoras')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/bitacoras') }}" class="nav-link"><i class="fas fa-clipboard-list"></i> Estaditicas del Mes </a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="panel shadow">

            <div class="header">
                <h2 class="title"><i class="fas fa-clipboard-list"></i> <strong>Estaditicas del Mes</strong></h2>

            </div>

            <div class="inside">
                <table id="table-modules" class="table table-striped table-hover">
                    <thead>
                        <tr id="fa"> 
                            <td rowspan="2" style="text-align:center;"><strong>SERVICIO</strong></td>
                            <td rowspan="2" style="text-align:center;"><strong>CITAS ATENDIDAS</strong></td>
                            <td rowspan="2" style="text-align:center;"><strong>EXAMENES REALIZADOS</strong></td>
                            <td colspan="9" style="text-align:center;"><strong>MATERIALES UTILIZADOS</strong></td>
                        </tr>
                        <tr>
                            <td style="text-align:center;"><strong>Placa 8x10</strong></td>
                            <td style="text-align:center;"><strong>Placa 10x12</strong></td>
                            <td style="text-align:center;"><strong>Placa 11x14</strong></td>
                            <td style="text-align:center;"><strong>Placa 14x17</strong></td>
                            <td style="text-align:center;"><strong>Fotos</strong></td>
                            <td style="text-align:center;"><strong>Hojas</strong></td>
                            <td style="text-align:center;"><strong>Pelicula 8x10</strong></td>
                            <td style="text-align:center;"><strong>Pelicula 10x12</strong></td>
                            <td style="text-align:center;"><strong>Singoplaza</strong></td>
                        </tr>
                    </thead>

                    <tbody>
                        @foreach($services as $ser)
                            <tr>
                                <td>{{ $ser->nombre }}</td>
                                <td>
                                    @foreach($citas_x_ser as $cser)
                                        @if($cser->id_service == $ser->id )
                                            {{ $cser->total_citas }}                                               
                                        @endif
                                    @endforeach
                                </td>
                                <td>
                                    @foreach($estudios_x_ser as $eser)
                                        @if($eser->id_service == $ser->id )
                                            {{ $eser->total_estudios }}                                               
                                        @endif
                                    @endforeach
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                                <td>
                                    
                                </td>
                            </tr>
                        @endforeach 
                        
                    </tbody>
                </table>
            </div>

        </div>
    </div>

@endsection