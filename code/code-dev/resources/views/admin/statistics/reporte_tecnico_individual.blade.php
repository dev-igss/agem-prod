<!DOCTYPE>
<html>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <link rel="shortcut icon" href="{{ url('/img/Isotipo.ico') }}" type="image/x-icon">
    <title> SPS-645</title>
    <style>
        table {
            border-collapse: collapse;
        }
        table,
        th,
        td {
            border: 1px solid black;
        }
        th,
        td {
            padding: 5px;
        }

        footer {
            position: fixed; 
            bottom: -30px; 
            left: 0px; 
            right: 0px;
            height: 30px; 
        }

        .celda{
            height: auto;
            width: 200px;
        }

        .celda1{
            height: auto;
            width: 140px;
        }
    </style>

    <body style="font-size: 14px; font-family: 'Roboto Slab', serif;">
        <footer>
            U=UTILES D=DESCARTABLES

            <div style="float: right; margin-top: 0px; font-size: 0.75em;">
                <span>Depto. Servicios de Apoyo IGSS  </span>
            </div>
        </footer>

        <div style="float: left; margin-top: -25px;" >
            <img src="{{ url('static/imagenes/igss.png') }}" alt="" width="80" height="80"/>
        </div>
        <div style="text-align: center; margin-top: 0px; margin-left: -800px;">
            <span><strong> SERVICIO DE RADIOLOGIA </strong> </span>
            <br>
            <span><strong>CONTROL DIARIO DE TRABAJO TECNICO INDIVIDUAL</strong></span>
        </div>
        <div style="float: right; margin-top: -25px;">
            <span><strong> SPS-645 </strong> </span>
        </div>
        <div style="float: left; margin-top: 40px; margin-left: -80px;  width: 500px; height: 25px; font-size: 16px;" >
            <span ><strong> Unidad Medica: </strong> IGSS Hospital General de Quetzaltenango </span>            
        </div>
        <div style="float: right; margin-top: 40px; margin-right: -200px;  width: 800px; height: 25px; font-size: 16px;" >
            <span ><strong> Nombre y Firma del Técnico en Rayos X: </strong> {{ $tecnico->name.' '.$tecnico->lastname }}  </span>                 
        </div>
        <br><br>
        <div style="float: left; margin-top: 40px; margin-left: -420px;  width: 500px; height: 25px; font-size: 16px;" >
            <span ><strong> Turno de: __________ a __________ horas </strong>  </span>            
        </div>
        <div style="float: right; margin-top: 40px; margin-right: -500px;  width: 500px; height: 25px; font-size: 16px;" >
            <span ><strong> Fecha: </strong> {{ $fecha }}   </span>                 
        </div>
        
        <br>
        <table width="100%"  style=" margin-top:60px; text-align: center; font-size: 12px;">
            <!-- Encabezado de la Tabla-->
            <tr>
                <th ROWSPAN="3">NOMBRE DEL PACIENTE</th>
                <th ROWSPAN="3">AFILIACION</th>
                <th ROWSPAN="3">ESTUDIO</th>
                <th COLSPAN="10">PLACAS RADIOGRAFICAS</th>
                <th ROWSPAN="3">EXP.</th>
                <th ROWSPAN="3" style="height: auto; width: 125px;">No. de Exposiciones</th>
                <th ROWSPAN="3" style="height: auto; width: 125px;">Persona que Recibe</th>              
            </tr>
            <tr>
                <th COLSPAN="2">8X10</th>
                <th COLSPAN="2">10X12</th>
                <th COLSPAN="2">11X14</th>
                <th COLSPAN="2">14X14</th>
                <th COLSPAN="2">14X17</th>
            </tr>
            <tr>
                <th>U</th>
                <th>D</th>
                <th>U</th>
                <th>D</th>
                <th>U</th>
                <th>D</th>
                <th>U</th>
                <th>D</th>
                <th>U</th>
                <th>D</th>
            </tr>

            <!-- Cuerpo de la Tabla-->
            @foreach($appointments as $a)
                <tr>
                    <th class="celda" >{{ $a->paciente }}</th>
                    <th >{{ $a->afiliacion }}</th>
                    <th class="celda">{{ $a->estudio }}</th>
                    <!-- 8x10 -->
                    <th > 
                        @if($a->material == '0') 
                            {{ $a->cantidad_material }}
                        @endif
                    </th>
                    <th ></th>
                    <!-- 10x12 -->
                    <th > 
                        @if($a->material == '1') 
                            {{ $a->cantidad_material }}
                        @endif
                    </th>
                    <th ></th>
                    <!-- 11x14 -->
                    <th > 
                        @if($a->material == '2') 
                            {{ $a->cantidad_material }}
                        @endif
                    </th>
                    <th ></th>
                    <!-- 14x14 -->
                    <th ></th>
                    <th ></th>
                    <!-- 14x17 -->
                    <th > 
                        @if($a->material == '3') 
                            {{ $a->cantidad_material }}
                        @endif
                    </th>                  
                    <th ></th>
                    <th >{{ $a->num_exp }}</th>
                    <th style="height: auto; width: 125px;"></th>
                    <th style="height: auto; width: 125px;"></th>
                </tr>
            @endforeach
        </table>

        
    </body>

    
</html>