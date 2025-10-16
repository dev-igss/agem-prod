@extends('admin.master')
@section('title','Dashboard')

@section('content')
    <div class="container-fluid">

        <div class="panel shadow">
            <div class="header">
                <h2 class="title"><i class="fas fa-chart-bar"></i><strong> Estadísticas Rápidas</strong></h2>
            </div>
        </div>

        <div class="row mtop16">

            <div class="col-md-4">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-users"></i><strong> Citas Agendadas</strong></h2>
                    </div>
                    <div class="inside">
                        <div class="big_count">
                            {{ $citas_d }}
                        </div>
                    </div>
                </div>
            </div>

            <div class="col-md-4">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-users"></i><strong> Citas en Atención</strong></h2>
                    </div>
                    <div class="inside">
                        <div class="big_count">
                            {{ $citas_d_aten }}
                        </div>
                    </div>
                </div>
            </div>

            <div class="col-md-4">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fas fa-users"></i><strong> Citas Ausentes</strong></h2>
                    </div>
                    <div class="inside">
                        <div class="big_count">
                            {{ $citas_d_aus }}
                        </div>
                    </div>
                </div>
            </div>

            
        </div>  

        <div class="panel shadow">
            <div class="header">
                <h2 class="title"><i class="fas fa-newspaper"></i><strong> Noticias o Alertas del Sistema - {{$today}}</strong></h2>
            </div>
        </div>

        <div class="row mtop16">

            <div class="col-md-12">
                <div class="panel shadow">
                    
                    <div class="inside">
                        <p style="font-size: 18px;">
                            <ul>
                                <li>
                                    Para un mejor control de los módulos de citas, fueron separadas. Verifique en el menu el módulo de citas que desea acceder
                                </li>
                                <li>
                                    El sistema toma en cuenta los dias festivos, la jefatura del departamento de Radiologia sera la encargada de asignar los mismos.
                                    Tomar en cuenta que las unicas citas validas para estos dias son de Hospitalización o Emergencia.
                                </li>
                                <li>
                                    Informar a los tecnicos que tienen problema de inicio de sesión, que intenten en esta nueva versión del sistema. En caso de seguir presentando
                                    problemas que secretaria o jefatura envie un correo a ricardod.velasquez@igssgt.org, con los datos del tecnico.
                                </li>
                            </ul>
                        </p>
                    </div>
                </div>
            </div>



            
        </div>  

    </div>

@endsection