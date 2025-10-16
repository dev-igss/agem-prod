@extends('admin.master')
@section('title','Calendario de Citas General')

@section('breadcrumb')
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/citas') }}" class="nav-link"><i class="fas fa-columns"></i> Citas</a>
    </li>
    <li class="breadcrumb-item">
        <a href="{{ url('/admin/equipments/add') }}" class="nav-link"><i class="fas fa-plus-circle"></i> Calendario de Citas General</a>
    </li>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-md-2">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-th-list"></i><strong> Clasificación de Citas</strong> </h2>                      
                    </div>                

                    <div class="inside">
                        <span><i class="fa fa-square" style="color: #1c2833;"></i> Cita RX</span><br>
                        <span><i class="fa fa-square" style="color: #f4d03f;"></i> Cita USG A.M</span><br>
                        <span><i class="fa fa-square" style="color: #27ae60;"></i> Cita USG P.M</span><br>
                        <span><i class="fa fa-square" style="color: #f5b7b1;"></i> Cita MMO</span><br>
                        <span><i class="fa fa-square" style="color: #e74c3c;"></i> Cita DMO</span><br>   
                    </div>  
                </div> 
            </div>

            <div class="col-md-10">
                <div class="panel shadow">
                    <div class="header">
                        <h2 class="title"><i class="fa fa-calendar"></i><strong> Calendario de Citas General</strong> </h2>
                        
                    </div>                

                    <div class="inside">
                        <div id="calendar"></div>                        
                    </div>  
                </div> 
            </div>
        </div>
    </div>

    <script>


        $(document).ready(function() {
            
            var currentLangCode = 'es';//cambiar el idioma al español
 
            $('#calendar').fullCalendar({
                
                    eventClick: function(calEvent, jsEvent, view) {
    
                        $(this).css('background', 'red');
                        //al evento click; al hacer clic sobre un evento cambiara de background
                        //a color rojo y nos enviara a los datos generales del evento seleccionado
                    },
                    initialView: "agendaDay",                
                    timeZone: "America/Guatemala",
                    header: {
                        left: 'prev,next today myCustomButton',
                        center: 'title',
                        right: 'month,agendaWeek,agendaDay'
                    },
                    buttonText: {
                        today: 'Hoy',
                        month: 'Mes',
                        week: 'Semana',
                        day: 'Día'
                    },
                    
                    monthNames: ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'],
                        monthNamesShort: ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'],
                        dayNames: ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'],
                        dayNamesShort: ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'],
                        
                    
    
                    lang:currentLangCode,
                    editable: true,
                    eventLimit: true,
                    
                    events:{
                        //para obtener los resultados del controlador y mostrarlos en el calendario
                        //basta con hacer referencia a la url que nos da dicho resultado, en el ejemplo
                        //en la propiedad url de events ponemos el enlace
                        //y listo eso es todo ya el plugin se encargara de acomodar los eventos
                        //segun la fecha.
                        //url:'http://agem.hgq/admin/agem/api/load/appointments',
                        url:'http://10.11.0.30:8900/admin/agem/api/load/appointments',
                        
                        
                    },
                    eventRender: function(event, element) {

                    if(event.area == "0") {
                        element.css('background-color', '#1c2833');
                        element.css('border-color', '#1c2833');
                        element.css('color', '#fff');
                    }
                    if(event.area == "2") {
                        if(event.type == "0"){
                            element.css('background-color', '#f4d03f');
                            element.css('border-color', '#f4d03f');
                            element.css('color', '#000');
                        }

                        if(event.type == "1"){
                            element.css('background-color', '#27ae60');
                            element.css('border-color', '#27ae60');
                            element.css('color', '#000');
                        }
                    }
                    if(event.area == "3") {
                        element.css('background-color', '#f5b7b1');
                        element.css('border-color', '#f5b7b1');
                        element.css('color', '#000');
                    }
                    if(event.area == "4") {
                        element.css('background-color', '#e74c3c');
                        element.css('border-color', '#e74c3c');
                        element.css('color', '#fff');
                    }
                },
            });
 
        });

        

        
    </script>
@endsection