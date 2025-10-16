<?php

use App\Http\Controllers\ProfileController;
use Illuminate\Support\Facades\Route;
use App\Http\Controllers\Auth\AuthenticatedSessionController;
use App\Http\Controllers\Admin\UserController;
use App\Http\Controllers\Admin\DashboardController;
use App\Http\Controllers\Admin\PatientController;
use App\Http\Controllers\Admin\ScheduleController;
use App\Http\Controllers\Admin\StudieController;
use App\Http\Controllers\Admin\ServiceController;
use App\Http\Controllers\Admin\UnitsController;
use App\Http\Controllers\Admin\ReportController;
use App\Http\Controllers\Admin\AppointmentController;
use App\Http\Controllers\Admin\BitacoraController;
use App\Http\Controllers\Admin\ApiController;
use App\Http\Controllers\ApiController as ApiControllerExternan;
use App\Http\Controllers\PatientDayController;

Route::get('/', [AuthenticatedSessionController::class, 'create'])->name('login');

Route::get('/dashboard', function () {
    return view('dashboard');
})->middleware(['auth', 'verified'])->name('dashboard');

Route::middleware('auth')->group(function () {
    Route::get('/profile', [ProfileController::class, 'edit'])->name('profile.edit');
    Route::patch('/profile', [ProfileController::class, 'update'])->name('profile.update');
    Route::delete('/profile', [ProfileController::class, 'destroy'])->name('profile.destroy');   
   
});

require __DIR__.'/auth.php';

Route::prefix('admin')->group(function () {
    //Dashboard
    Route::get('/dashboard', [DashboardController::class, 'getDashboard'])->name('admin_dashboard');

    Route::get('/bitacoras', [BitacoraController::class, 'getBitacora'])->name('bitacoras');
    Route::post('/bitacora/busqueda', [BitacoraController::class, 'postSearch'])->name('bitacoras');
    Route::post('/bitacora/busqueda/', [BitacoraController::class, 'postSearch'])->name('bitacoras');

    //Citas
    Route::get('/citas/rx', [AppointmentController::class, 'getAppointmentRx'])->name('appointment_rx');
    Route::get('/citas/umd', [AppointmentController::class, 'getAppointmentUmd'])->name('appointment_umd');
    Route::get('/cita/agregar', [AppointmentController::class, 'getAppointmentAdd'])->name('appointment_add');        
    Route::post('/cita/agregar', [AppointmentController::class, 'postAppointmentAdd'])->name('appointment_add');
    Route::post('/cita/busqueda', [AppointmentController::class, 'postAppointmentSearch'])->name('appointment_list');
    Route::get('/cita/calendario', [AppointmentController::class, 'getCalendar'])->name('appointment_calendar'); 
    Route::get('/cita/calendario/rx', [AppointmentController::class, 'getCalendarRx'])->name('appointment_calendar'); 
    Route::get('/cita/calendario/umd', [AppointmentController::class, 'getCalendarUmd'])->name('appointment_calendar'); 
    Route::get('/cita/{id}/materiales', [AppointmentController::class, 'getAppointmentMaterials'])->name('appointment_materials');
    Route::get('/cita/materiales/{id}/material_desechado', [AppointmentController::class, 'getAppointmentMaterialsDiscarded'])->name('appointment_materials');
    Route::get('/cita/{id}/registro_materiales/{idstudy}/{idmaterial}/{amount}', [AppointmentController::class, 'getAppointmentRegisterMaterials'])->name('appointment_materials');
    Route::get('/cita/{id}/reprogramar/{date}', [AppointmentController::class, 'getAppointmentReschedule'])->name('appointment_reschedule');
    Route::get('/cita/{id}/cambio_horario/{horario}', [AppointmentController::class, 'getScheduleChange'])->name('appointment_reschedule');
    Route::get('/cita/{id}/reprogramacion_forzada/{date}/{comment?}', [AppointmentController::class, 'getAppointmentReschedule'])->name('appointment_reschedule'); 
    Route::get('/cita/{id}/paciente_presente/{status}', [AppointmentController::class, 'getAppointmentPatientsStatus'])->name('appointment_patients_status');
    Route::get('/cita/{id}/paciente_ausente/{status}', [AppointmentController::class, 'getAppointmentPatientsStatus'])->name('appointment_patients_status');
    Route::get('/cita/{id}/informe_al_patrono', [AppointmentController::class, 'getAppointmentInforme'])->name('appointment_materials');
    Route::get('/cita/configuracion', [AppointmentController::class, 'getConfigAppointment'])->name('appointment_setting');
    Route::post('/cita/configuracion', [AppointmentController::class, 'postConfigAppointment'])->name('appointment_setting');
    Route::get('/cita/configuracion/dias/festivos', [AppointmentController::class, 'getConfigHolyDays'])->name('appointment_setting');
    Route::post('/cita/configuracion/dias/festivos', [AppointmentController::class, 'postConfigHolyDays'])->name('appointment_setting');    
    Route::get('/cita/configuracion/dias/festivos/{id}/borrar', [AppointmentController::class, 'getSettingsHolyDaysDelete'])->name('appointment_setting');
    Route::get('/cita/{id}/borrar', [AppointmentController::class, 'getAppointmentDelete'])->name('appointment_delete');

    //Units
    Route::get('/unidades', [UnitsController::class, 'getHome'])->name('units');
    Route::post('/unidad/agregar', [UnitsController::class, 'postUnitAdd'])->name('unit_add');
    Route::get('/unidad/{id}/editar', [UnitsController::class, 'getUnitEdit'])->name('unit_edit');
    Route::post('/unidad/{id}/editar', [UnitsController::class, 'postUnitEdit'])->name('unit_edit');
    Route::get('/unidad/{id}/borrar', [UnitsController::class, 'getUnitDelete'])->name('unit_delete');

    //Services
    Route::get('/servicios_general', [ServiceController::class,'getHome'])->name('serviceg_list');
    Route::post('/servicios_general/agregar', [ServiceController::class,'postServicesGeneralAdd'])->name('serviceg_add');        
    Route::get('/servicios_general/{id}/editar', [ServiceController::class,'getServicesGeneralEdit'])->name('serviceg_edit');
    Route::post('/servicios_general/{id}/editar', [ServiceController::class,'postServicesGeneralEdit'])->name('serviceg_edit');
    Route::get('/servicios_general/{id}/servicios',[ServiceController::class,'getServicesGeneralServices'])->name('service_list');
    Route::post('/servicios_general/servicios/agregar',[ServiceController::class,'postServicesGeneralServicesAdd'])->name('service_add');
    Route::get('/servicios_general/servicios/{id}/editar',[ServiceController::class,'getServicesGeneralServicesEdit'])->name('service_edit');
    Route::post('/servicios_general/servicios/{id}/editar',[ServiceController::class,'postServicesGeneralServicesEdit'])->name('service_edit');

    //Studies
    Route::post('/estudio/agregar', [StudieController::class,'postStudieAdd'])->name('studie_add'); 
    Route::get('/estudios/{type}', [StudieController::class,'getHome'])->name('studie_list');       
    Route::get('/estudio/{id}/editar', [StudieController::class,'getStudieEdit'])->name('studie_edit');
    Route::post('/estudio/{id}/editar', [StudieController::class,'postStudieEdit'])->name('studie_edit');

    //Horarios
    Route::get('/horarios', [ScheduleController::class,'getHome'])->name('schedule_list');
    Route::post('/horario/agregar', [ScheduleController::class,'postScheduleAdd'])->name('schedule_add');
    Route::get('/horario/{id}/editar', [ScheduleController::class,'getScheduleEdit'])->name('schedule_edit');
    Route::post('/horario/{id}/editar', [ScheduleController::class,'postScheduleEdit'])->name('schedule_edit');

    //Patients        
    Route::get('/paciente/agregar', [PatientController::class,'getPatientAdd'])->name('patient_add');        
    Route::post('/paciente/agregar', [PatientController::class,'postPatientAdd'])->name('patient_add');   
    Route::get('/pacientes/{filtro}', [PatientController::class,'getHome'])->name('patient_list');  
    Route::post('/paciente/busqueda', [PatientController::class,'postSearch'])->name('patient_list');       
    Route::get('/paciente/{id}/editar', [PatientController::class,'getPatientEdit'])->name('patient_edit');
    Route::post('/paciente/{id}/editar', [PatientController::class,'postPatientEdit'])->name('patient_edit');        
    Route::get('/paciente/configuracion', [PatientController::class,'getConfigPatient'])->name('patient_setting');
    Route::post('/paciente/configuracion', [PatientController::class,'postConfigPatient'])->name('patient_setting');
    Route::get('/paciente/{id}/historial_citas', [PatientController::class,'getPatientHistoryExam'])->name('patient_history_exam');
    Route::get('/paciente/{id}/historial_codigos_expedientes', [PatientController::class,'getPatientHistoryCode'])->name('patient_history_exam');
    Route::get('/paciente/{id}/borrar', [PatientController::class,'getPatientDelete'])->name('patient_delete');
    Route::get('/paciente/{id}/restablecer', [PatientController::class,'getPatientRestore'])->name('patient_restore');

    //Reportes
    Route::get('/reportes',[ReportController::class, 'getHome'])->name('reports');      
    Route::post('/reporte/filtrado/fechas',[ReportController::class, 'postStaticsBetweenDates'])->name('reports');  
    Route::get('/reporte/mensual/citas',[ReportController::class, 'getStaticsMonth'])->name('report_month_areas');  
    Route::post('/reporte/estadisticas/dmo',[ReportController::class, 'postReportDMOEstadistica'])->name('report_month_areas');
    Route::post('/reporte/estadisticas/mamo',[ReportController::class, 'postReportMAMOEstadistica'])->name('report_month_areas');
    Route::post('/reporte/estadisticas/usg',[ReportController::class, 'postReportUSGEstadistica'])->name('report_month_areas');
    Route::post('/reporte/estadisticas/rx',[ReportController::class, 'postReportRXEstadistica'])->name('report_month_areas');
    Route::post('/reporte/tecnicos/individual',[ReportController::class, 'postReportTecnicoIndividual'])->name('report_tecnicos_individual');

    //Users
    Route::get('/usuarios', [UserController::class,'getUsers'])->name('user_list');
    Route::post('/usuario/agregar', [UserController::class,'postUserAdd'])->name('user_add');
    Route::get('/usuario/{id}/editar', [UserController::class,'getUserEdit'])->name('user_edit');
    Route::post('/usuario/{id}/editar', [UserController::class,'postUserEdit'])->name('user_edit');
    Route::post('/usuario/{id}/reiniciar_contrasena',[UserController::class,'postResetPassword'])->name('user_reset_password');
    Route::get('/usuario/{id}/suspender', [UserController::class,'getUserBanned'])->name('user_banned');
    Route::get('/usuario/{id}/permisos', [UserController::class,'getUserPermissions'])->name('user_permissions');
    Route::post('/usuario/{id}/permisos', [UserController::class,'postUserPermissions'])->name('user_permissions');
    Route::get('/usuario/{id}/solicitudes_fuera_de_tiempo', [UserController::class,'getUserRequestsOut'])->name('user_requests_out');
    Route::post('/usuario/{id}/solicitudes_fuera_de_tiempo', [UserController::class,'postUserRequestsOut'])->name('user_requests_out');
    Route::get('/usuario/cuenta/informacion',[UserController::class,'getAccountInfo'])->name('user_info');
    Route::post('/usuario/cuenta/cambiar/contrasena',[UserController::class,'postAccountChangePassword'])->name('user_change_password');

    //Request Ajax 
    Route::get('/agem/api/load/add/patient/{code}/{exam}', [ApiController::class,'getPatient']);
    Route::get('/agem/api/load/add/patient/beneficiario/{code}/{exam}', [ApiController::class,'getPatientBeneficiario']);
    Route::get('/agem/api/load/generate/code/{code}', [ApiController::class,'getCodePatient']);
    Route::get('/agem/api/load/studies/{type}', [ApiController::class,'getStudies']);
    Route::get('/agem/api/load/appointments/{date}/{area}', [ApiController::class,'getAppointments']);
    Route::get('/agem/api/load/schedules/{date}/{area}', [ApiController::class,'getSchedule']);
    Route::get('/agem/api/load/control/studies/{date}', [ApiController::class,'getStudiesControlDate']);
    Route::get('/agem/api/load/holy/days/{date}', [ApiController::class,'getHolyDays']);
    Route::get('/agem/api/load/schedules/change', [ApiController::class,'getScheduleChange']);
    Route::get('/agem/api/load/appointments', [ApiController::class,'getAppointmentsView']);
    Route::get('/agem/api/load/appointments/rx', [ApiController::class,'getAppointmentsViewRx']);
    Route::get('/agem/api/load/appointments/umd', [ApiController::class,'getAppointmentsViewUmd']);
    Route::get('/agem/api/load/services', [ApiController::class,'getServices']);
    Route::get('/agem/api/load/consulta/medicos/{mes}/{year}', [ApiController::class,'getPruebaConsulta']);
    Route::get('/agem/api/load/consulta/dias/festivos/{year}', [ApiController::class,'getHolyDays']);

    Route::get('/agem/api/load/consulta/afiliacion/medi', [ApiController::class,'pruebaConsultaApi']);

});

Route::get('/citas_del_dia_rx', [PatientDayController::class,'getPatientDayRx'])->name('patient_day');
Route::get('/citas_del_dia_umd/{filtrado}', [PatientDayController::class,'getPatientDayUmd'])->name('patient_day');
Route::get('/citas_abiertas/{filtrado}', [PatientDayController::class,'getPatientDayOpen'])->name('patient_day');
Route::get('/citas_del_dia/{id}/materiales', [PatientDayController::class,'getMaterials'])->name('materials');
Route::post('/citas_del_dia/materiales', [PatientDayController::class,'postMaterials'])->name('materials');  
Route::get('/citas_del_dia/acciones/{id}/comentario/{text}', [PatientDayController::class,'getAppointmentComment'])->name('materials'); 
Route::get('/citas_del_dia/acciones/{id}/solicitud_reprogramacion', [PatientDayController::class,'getAppointmentReschedule'])->name('materials'); 
Route::get('/citas_del_dia/acciones/{id}/ausente_examen', [PatientDayController::class,'getAppointmentNot'])->name('materials'); 
Route::get('/citas_del_dia/acciones/{id}/agregar_estudio/{area}/{study}/{comment}', [PatientDayController::class,'getAppointmentAddExamen'])->name('materials'); 

//Request Ajax
Route::get('/agem/api/load/name/study/{id}', [ApiControllerExternan::class,'getStudyName']);
Route::get('/agem/api/load/name/study/all/{type}', [ApiControllerExternan::class,'getStudy']);
