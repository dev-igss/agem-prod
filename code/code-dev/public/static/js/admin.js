const http = new XMLHttpRequest();
const csrfToken = document.getElementsByName('csrf-token')[0].getAttribute('content');
var base = location.protocol+'//'+location.host;
var route = document.getElementsByName('routeName')[0].getAttribute('content');

document.addEventListener('DOMContentLoaded', function(){
    var btn_search = document.getElementById('btn_search');
    var form_search = document.getElementById('form_search');

    if(btn_search){
        btn_search.addEventListener('click', function(e){
            e.preventDefault();
            if(form_search.style.display === 'block'){
                form_search.style.display = 'none';
            }else{
                form_search.style.display = 'block';
            }
        });
    }



    if(route == "product_edit"){
        var btn_product_file_image = document.getElementById('btn_product_file_image');
        var product_file_image = document.getElementById('product_file_image');

        btn_product_file_image.addEventListener('click', function(){
            product_file_image.click();
        }, false);

        product_file_image.addEventListener('change', function(){
            document.getElementById('form_product_gallery').submit();
        });
    }

    if(route == "diet_request_add"){
        getDietasDisponibles();
    }

    document.getElementsByClassName('lk-'+route)[0].classList.add('active');

    btn_deleted = document.getElementsByClassName('btn-deleted');
    for(i=0; i < btn_deleted.length; i++){
        btn_deleted[i].addEventListener('click', delete_object);
    }

    $('#table-modules').DataTable({
        "paging": true,
        "lengthChange": true,
        "searching": true,
        "ordering": true,
        "info": true,
        "autoWidth": false,
        "responsive": true,
        language: {
            "decimal": "",
            "emptyTable": "No hay registros",
            "info": "Mostrando _START_ a _END_ de _TOTAL_ Registros",
            "infoEmpty": "Mostrando 0 to 0 of 0 Registros",
            "infoFiltered": "(Filtrado de _MAX_ total registros)",
            "infoPostFix": "",
            "thousands": ",",
            "lengthMenu": "Mostrar _MENU_ Registros",
            "loadingRecords": "Cargando...",
            "processing": "Procesando...",
            "search": "Buscar:",
            "zeroRecords": "Sin resultados encontrados",
            "paginate": {
                "first": "Primero",
                "last": "Ultimo",
                "next": "Siguiente",
                "previous": "Anterior"
            }
        }
    });

    $("#idsupplier").select2({
        placeholder: "Seleccione una Opción",
        allowClear: true
    });

    $("#idproduct").select2({
        placeholder: "Seleccione una Opción",
        allowClear: true
    });

    $("#idservice").select2({
        placeholder: "Seleccione una Opción",
        allowClear: true
    });

    $("#pidarticulo").select2({
        placeholder: "Seleccione una Opción",
        allowClear: true
    });




});

function delete_object(e){
    e.preventDefault();
    var object = this.getAttribute('data-object');
    var action = this.getAttribute('data-action');
    var path = this.getAttribute('data-path');
    var nombre_servicio = this.getAttribute('data-servicio'); 
    var dietas_solicitadas= this.getAttribute('data-dietas'); 
    var dietas_servidas = this.getAttribute('data-dietas-served'); 

    

    if(action == "impresion_lote" || action == "impresion_cocineta"){
        var url = base + '/' + path + '/' + action;
    }else{
        var url = base + '/' + path + '/' + object + '/' + action;
    }

    var title, text, icon, cantidad_dietas;

    //console.log(url);

    if(action == "delete"){
        title = "¿Esta seguro de eliminar este elemento?";
        text = "Recuerde que esta acción enviara este elemento a la papelera o lo eliminara de forma definitiva.";
        icon = "warning";
    }

    if(action == "restore"){
        title = "¿Quiere restaurar este elemento?";
        text = "Esta acción restaurará este elemento y estará activo en la base de datos.";
        icon = "info";
    }

    if(action == "servida"){
        title = "¿Esta seguro de marcar como servida esta dieta?";
        text = "Recuerde verificar la cantidad de dietas a servir, antes de guardar.";
        icon = "warning";
    }

    if(action == "cambio_dietas_servidas"){
        title = "¿Esta seguro de cambiar la cantidad de dietas servidas esta solicitud?";
        text = "Recuerde verificar la cantidad de dietas a servir, antes de guardar.";
        icon = "warning";
    }

    if(action == "impresion_lote"){
        title = "¡Impresion de Solicitudes de Dietas por Lote!";
        text = "Seleccione la jornda del dia que desea imprimir por lote.";
        icon = "info";
    }

    if(action == "impresion_cocineta"){
        title = "¡Impresion de Solicitudes por Cocinetas!";
        text = "Seleccione la jornada del dia que desea imprimir por cocineta.";
        icon = "info";
    }

    if(action == "anular"){
        title = '¿Esta seguro de '+'"Anular"'+' está solicitud?';
        text = "Recuerde que esta acción enviara está solicitud a la papelera.";
        icon = "warning";
    }

    if(action == "servida"){
        Swal.fire({
            title: title,
            text: text,
            icon: icon,
            showCancelButton: true,
            confirmButtonText: 'Aceptar',
        }).then((result) =>{
            if (result.isConfirmed) {
                Swal.fire({
                    title: "Ingrese la cantidad de dietas a servir en esta solicitud <br/><br/>Servicio: "+nombre_servicio+"<br/>Diestas Solicitadas: "+dietas_solicitadas,
                    icon: "info",
                    confirmButtonText: 'Aceptar',
                    showDenyButton: true,
                    denyButtonText: 'Cancelar',
                    focusConfirm:true,
                    html:
                    '<label> <strong> Cantidad de dietas </strong> </label>'
                    +
                   '<br/><input type="number" min="0" id="swal-input3" class="swal2-input" style="width:100%;">',
                   preConfirm: () => {
                        return [
                            cantidad_dietas = document.getElementById('swal-input3').value
                        ]
                    }
                }).then((result) =>{
                    //console.log(cantidad_dietas);
                    if (result.isConfirmed && cantidad_dietas.length === 0) {

                        window.alert("Ingrese una cantidad de dietas");

                    }
                    
                    if (result.isConfirmed && cantidad_dietas.length > 0){
                        window.location.href = url+"/"+cantidad_dietas;
                        //console.log(url+"/"+cantidad_dietas);
                    }

                    if (result.isDenied) {    
                        window.alert("No se registraran dietas servidas");
                    }


                });
            }
        });

    }    else if(action == "cambio_dietas_servidas"){
        Swal.fire({
            title: title,
            text: text,
            icon: icon,
            showCancelButton: true,
            confirmButtonText: 'Aceptar',
        }).then((result) =>{
            if (result.isConfirmed) {
                Swal.fire({
                    title: "Ingrese la cantidad de dietas a servir en esta solicitud <br/>Dietas Ya Servidas: "+dietas_servidas,
                    icon: "info",
                    confirmButtonText: 'Aceptar',
                    showDenyButton: true,
                    denyButtonText: 'Cancelar',
                    focusConfirm:true,
                    html:
                    '<label> <strong> Cantidad de dietas </strong> </label>'
                    +
                   '<br/><input type="number" min="0" id="swal-input3" class="swal2-input" style="width:100%;">',
                   preConfirm: () => {
                        return [
                            cantidad_dietas = document.getElementById('swal-input3').value
                        ]
                    }
                }).then((result) =>{
                    //console.log(cantidad_dietas);
                    if (result.isConfirmed && cantidad_dietas.length === 0) {

                        window.alert("Ingrese una cantidad de dietas");

                    }
                    
                    if (result.isConfirmed && cantidad_dietas.length > 0){
                        window.location.href = url+"/"+cantidad_dietas;
                        //console.log(url+"/"+cantidad_dietas);
                    }

                    if (result.isDenied) {    
                        window.alert("No se registraran dietas servidas");
                    }


                });
            }
        });

    }
    
    else if(action == "impresion_lote"){
        Swal.fire({
            title: title,
            text: text,
            icon: icon,
            confirmButtonText: 'Aceptar',
            cancelButtonText: 'Cancelar',
            showCancelButton: true,
            html: '<select id="swal-input2" class="swal2-input"> <option value="1">Desayuno</option> <option value="2">Almuerzo</option> <option value="3">Cena</option> <option value="4">Refacción</option> </select>',
            focusConfirm: false,
            allowOutsideClick: false,
            preConfirm: () => {

                return document.getElementById('swal-input2').value;
            }
        }).then((result) =>{
            if (result.isConfirmed) {
                jornada = result.value;

                window.location.href = url+'/'+jornada;

                //console.log(url+'/'+jornada);
            }
        });

    }else if(action == "impresion_cocineta"){
        Swal.fire({
            title: title,
            text: text,
            icon: icon,
            confirmButtonText: 'Aceptar',
            cancelButtonText: 'Cancelar',
            showCancelButton: true,
            html: '<select id="swal-input2" class="swal2-input"> <option value="1">Desayuno</option> <option value="2">Almuerzo</option> <option value="3">Cena</option> </select>',
            focusConfirm: false,
            allowOutsideClick: false,
            preConfirm: () => {

                return document.getElementById('swal-input2').value;
            }
        }).then((result) =>{
            if (result.isConfirmed) {
                jornada = result.value;

                window.location.href = url+'/'+jornada;

                //console.log(url+'/'+jornada);
            }
        });

    }else if(action == "anular"){
        Swal.fire({
            title: title,
            text: text,
            icon: icon,
            showDenyButton: true,
            confirmButtonText: 'Aceptar',
            denyButtonText: 'Cancelar',
        }).then((result) =>{

            if (result.isConfirmed) {
                //console.log(url);
                window.location.href =  url;
            }
        });
    }else{
        Swal.fire({
            title: title,
            text: text,
            icon: icon,
            showCancelButton: true,
        }).then((result) =>{
            if (result.isConfirmed) {
                window.location.href = url;
            }
        });
    }


}

function getDietasDisponibles(){
    
    var inputdate = document.getElementById('idservice');
    var options1=$('#pidarticulo option').clone();

    console.log(options1);
    

    $(inputdate).change(function(){
        //console.log(inputdate.value);

        
            var url = base + '/admin/diat/api/load/diets';
            http.open('GET', url, true);
            http.setRequestHeader('X-CSRF-TOKEN', csrfToken);
            http.send();
            http.onreadystatechange = function(){
                if(this.readyState == 4 && this.status == 200){
                    var data = this.responseText;
                    data = JSON.parse(data);
                    //console.log(data);  
                    if(data.length > 0){
                        data.forEach( function(dietas, index){    
                            if(inputdate.value == 11 || inputdate.value == 16 || inputdate.value == 21 || inputdate.value == 57 || inputdate.value == 65 ){      
                                $('#pidarticulo').html(options1);                          
                                
                            }else{
                                $('#pidarticulo option[value="'+dietas.id+'"]').remove();  
                            }
                            //console.log(schedule.schedule_id);                                        
                        }); 
                    
                    }
                    
                    
                }
            } 
        
                
    });

    
}

