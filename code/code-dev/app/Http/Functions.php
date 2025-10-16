<?php
    //Rol de usuarios
    function getRoleUserArray($mode, $id){
        $roles = [
            '0' => 'Administrador General',
            '1' => 'Administrador',
            '2' => 'Encargado de Área',
            '3' => 'Técnico',
            '4' => 'Secretaria',
            '5' => 'Medico'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }

    //Estado de Usuarios
    function getUserStatusArray($mode, $id){
        $status = [
            '0' => 'Suspendido',
            '1' => 'Activo'
        ];

        if(!is_null($mode)):
            return $status;
        else:
            return $status[$id];
        endif;
    }

    function getGenderPatient($mode, $id){
        $roles = [
            '0' => 'Masculino',
            '1' => 'Femenino'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }

    function getHourType($mode, $id){
        $roles = [
            '0' => 'A.M',
            '1' => 'P.M'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }

    function getTypeStudie($mode, $id){
        $roles = [
            '0' => 'RX',
            '1' => 'RX Especiales',
            '2' => 'USG',
            '3' => 'MMO',
            '4' => 'DMO' 
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif; 
    }

    function getTypePatient($mode, $id){
        $roles = [
            '0' => 'AF',
            '1' => 'BE',
            '2' => 'BH',
            '3' => 'JB',            
            '4' => 'NA',
            '5' => 'PI',
            '6' => 'CS'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }

    function getExamB($mode, $id){
        $roles = [
            '0' => 'RX',
            '1' => 'RX Especial',
            '2' => 'USG',
            '3' => 'MMO',
            '4' => 'DMO'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }

    function getStatusAppointment($mode, $id){
        $roles = [
            '0' => 'Agendada',
            '1' => 'En Atención',
            '2' => 'Ausente a la Cita',
            '3' => 'Finalizada',
            '4' => 'Reprogramada',
            '5' => 'Solicitud de Reprogramación',
            '6' => 'Ausente al Examen'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }

    function getTypeAppointment($mode, $id){
        $roles = [
            '0' => 'Nueva Consulta',
            '1' => 'Reconsulta'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }

    function getMaterials($mode, $id){
        $roles = [
            '0' => 'Placa 8x10',
            '1' => 'Placa 10x12',
            '2' => 'Placa 11x14',
            '3' => 'Placa 14x17',            
            '4' => 'Hojas',
            '5' => 'Fotos',
            '6' => 'Pelicula 8x10',
            '7' => 'Pelicula 10x12',
            '8' => 'Singoplaza'
        ];

        if(!is_null($mode)):
            return $roles;
        else:
            return $roles[$id];
        endif;
    }


    //Key Value From JSON
    function kvfj($json, $key){
        if($json == null):
            return null;
        else:
            $json = $json;
            $json = json_decode($json, true);

            if(array_key_exists($key, $json)):
                return $json[$key];
            else:
                return null;
            endif;
        endif;
    }

    function user_permissions(){
        $p = [

            'dashboard' => [
                'icon' => '<i class="fas fa-tachometer-alt"></i>',
                'title' => 'Modulo de Dashboard',
                'keys' => [
                    'dashboard' => 'Puede ver el dashboard.',
                    'dashboard_small_stats' => 'Puede ver las estadísticas rápidas del día.',
                    'dashboard_small_tec_ser' => 'Puede ver las estadísticas por tecnicos y servicios.',
                    'dashboard_small_stats_month' => 'Puede ver las estadísticas rápidas del mes.',
                    'dashboard_prod_tecnicos_stats' => 'Puede ver las estadísticas de productividad de todos los tecnicos por area.',
                    'dashboard_prod_individual_stats' => 'Puede ver las estadísticas de productividad individual por area.',
                    'dashboard_appointments_tec' => 'Puede ver las citas atendidas por el tecnico.'
                ]
            ],

            'units' => [
                'icon' => '<i class="fas fa-hospital-user"></i>',
                'title' => 'Modulo de Unidades Medicas',
                'keys' => [
                    'units' => 'Puede ver el listado de unidades.',
                    'unit_add' => 'Puede agregar nuevas unidades.',
                    'unit_edit' => 'Puede editar unidades.',
                    'unit_delete' => 'Puede eliminar unidades.',
                    'unit_search' => 'Puede buscar unidades.'
                ]
            ],

            'bitacoras' => [
                'icon' => '<i class="fas fa-clipboard-list"></i> ',
                'title' => 'Modulo de Bitacoras',
                'keys' => [
                    'bitacoras' => 'Puede ver el listado de bitacoras.'
                ]
            ],

            'users' => [
                'icon' => '<i class="fas fa-user-lock"></i> ',
                'title' => 'Modulo de Usuarios',
                'keys' => [
                    'user_list' => 'Puede ver el listado de usuarios.',
                    'user_add' => 'Puede agregar nuevos usuarios.',
                    'user_edit' => 'Puede editar usuarios.',
                    'user_banned' => 'Puede suspender usuarios.',
                    'user_delete' => 'Puede eliminar usuarios.',
                    'user_reset_password' => 'Puede restablecer contraseña de usuarios.',
                    'user_permissions' => 'Puede administrar los permisos de los usuarios.',
                    'user_info' => 'Puede ver información de su cuenta',
                    'user_change_password' => 'Puede cambiar su contraseña de inicio de sesión'
                ]
            ],

            'services' => [
                'icon' => '<i class="fas fa-object-group"></i> ',
                'title' => 'Modulo de Servicios',
                'keys' => [
                    'serviceg_list' => 'Puede ver el listado de servicios general.',
                    'serviceg_add' => 'Puede agregar servicios genereales.',
                    'serviceg_edit' => 'Puede editar servicios generales.',
                    'service_list' => 'Puede ver el listado de servicios .',
                    'service_add' => 'Puede agregar servicios.',
                    'service_edit' => 'Puede editar servicios.'
                ]
            ],

            'studies' => [
                'icon' => '<i class="fas fa-heartbeat"></i> ',
                'title' => 'Modulo de Examenes / Estudios',
                'keys' => [
                    'studie_list' => 'Puede ver el listado de examenes.',
                    'studie_add' => 'Puede agregar examenes.',
                    'studie_edit' => 'Puede editar examenes.'
                ]
            ],
            
            'patients' => [
                'icon' => '<i class="fas fa-users"></i> ',
                'title' => 'Modulo de Pacientes',
                'keys' => [
                    'patient_list' => 'Puede ver el listado de pacientes.',
                    'patient_add' => 'Puede agregar pacientes.',
                    'patient_edit' => 'Puede editar pacientes.',
                    'patient_edit_affiliation' => 'Puede editar la afiliación de los pacientes.',
                    'patient_delete' => 'Puede borrar pacientes.',
                    'patient_restore' => 'Puede restaurar pacientes.',
                    'patient_history_exam' => 'Puede ver el historial de citas de pacientes.',
                    'patient_setting' => 'Puede modificar la configuración de los correlativos de expedientes.'
                ]
            ],

            'schedules' => [
                'icon' => '<i class="fa fa-clock-o"></i> ',
                'title' => 'Modulo de Horarios',
                'keys' => [
                    'schedule_list' => 'Puede ver el listado de horarios.',
                    'schedule_add' => 'Puede agregar horarios.',
                    'schedule_edit' => 'Puede editar horarios.'
                ]
            ],

            'appointments' => [
                'icon' => '<i class="fas fa-calendar-alt "></i>',
                'title' => 'Modulo de Citas',
                'keys' => [
                    'appointment_list' => 'Puede ver el listado de citas.',
                    'appointment_calendar' => 'Puede ver el calendario de citas.',
                    'appointment_add' => 'Puede agendar citas.',
                    'appointment_materials' => 'Puede reagendar citas a pacientes.',
                    'appointment_reschedule' => 'Puede ver el listado de materiales utilizados.',
                    'appointment_patients_status' => 'Puede marcar como presentes o ausentes a los pacientes.',
                    'appointment_setting' => 'Puede modificar la configuración de las citas de todo el servicio.',
                    'appointment_delete' => 'Puede eliminar citas agendadas.'                  
                          
                ]
            ],

            'reports' => [
                'icon' => '<i class="fa-solid fa-chart-column"></i>',
                'title' => 'Modulo de Reportes',
                'keys' => [
                    'reports' => 'Puede ver el modulo de reportes.',
                    'report_dates' => 'Puede generar reporte rapido entre fechas.',
                    'report_month_areas' => 'Puede generar el reporte mensual de cada área.',             
                    'report_tecnicos_individual' => 'Puede generar el reporte SPS-645 por técnico'      
                ] 
            ]
            

        ];

        return $p;
    }

    function getUserYears(){
        $ya = date('Y');
        $ym = $ya - 18;
        $yo = $ym - 62;

        return [$ym, $yo];
    }

    function getMonths($mode, $key){
        $m = [
            '01' => 'Enero',
            '02' => 'Febrero',
            '03' => 'Marzo',
            '04' => 'Abril',
            '05' => 'Mayo',
            '06' => 'Junio',
            '07' => 'Julio',
            '08' => 'Agosto',
            '09' => 'Septiembre',
            '10' => 'Octubre',
            '11' => 'Noviembre',
            '12' => 'Diciembre'
        ];

        if($mode == "list"){
            return $m;
        }else{
            return $m[$key];
        }
    }

    function number($number){
        return 'Q'.number_format($number, 2, '.', ',' );
    }

?>
