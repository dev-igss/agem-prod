const base = location.protocol+'//'+location.host;
const route = document.getElementsByName('routeName')[0].getAttribute('content');
const http = new XMLHttpRequest();
const csrfToken = document.getElementsByName('csrf-token')[0].getAttribute('content');
const PBX  = "PBX"; 
var page = 1;
var page_section = "";

var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'))
var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
  return new bootstrap.Tooltip(tooltipTriggerEl)
})


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
    
    var slider = new MDSlider;
    var form_avatar_change = document.getElementById('form_avatar_change');
    var btn_avatar_edit = document.getElementById('btn_avatar_edit');
    var avatar_change_overlay = document.getElementById('avatar_change_overlay');
    var input_file_avatar = document.getElementById('input_file_avatar');
    var units_list = document.getElementById('units_list');
    var load_more_units = document.getElementById('load_more_units');

    if(load_more_units){
        load_more_units.addEventListener('click', function(e){
            e.preventDefault();
            load_units(page_section);
        })
    }

    if(btn_avatar_edit){
        btn_avatar_edit.addEventListener('click', function(e){
            e.preventDefault();
            input_file_avatar.click();
        })
    }

    if(input_file_avatar){
        input_file_avatar.addEventListener('change', function(){
            var load_img = '<img src="'+base+'/static/imagenes/loader_white.svg"/>';
            avatar_change_overlay.innerHTML = load_img;
            avatar_change_overlay.style.display = 'flex';
            form_avatar_change.submit();
        })
    }

    slider.show();

    if(route == "home"){
        load_units('home');
    }

});

function load_units(section){
    page_section = section;
    var url = base + '/md/api/load/units/'+page_section+'?page='+page;
    http.open('GET', url, true);
    http.setRequestHeader('X-CSRF-TOKEN', csrfToken);
    http.send();
    http.onreadystatechange = function(){
        if(this.readyState == 4 && this.status == 200){
            page = page + 1;
            var data = this.responseText;
            data = JSON.parse(data);

            if(data.data.length == 0){
                load_more_units.style.display = "none";
            }

            data.data.forEach( function(unit, index){
                var div = "";
                div += "<div class=\"unit\">";                    
                    div += "<div class=\"image\">"
                        div += "<div class=\"overlay\">";
                            div += "<div class=\"btns\">";
                                div += "<a href=\""+base+"/unit/"+unit.id+"/services/todos"+"\"><i class=\"fas fa-eye\"></i></a>";
                            div += "</div>";
                        div += "</div>";
                        div += "<img src=\""+base+"/uploads/"+unit.file_path+"/"+unit.image+"\">";
                    div += "</div>";
                    div += "<a href=\""+base+"/unit/"+unit.id+"/services/todos"+"\">";
                        div += "<div class=\"title\">"+unit.name+"</div>";
                        div += "<div class=\"price\">"+PBX+": "+unit.phone+"</div>";
                        div += "<div class=\"options\"></div>";
                    div += "</a>";
                div += "</div>";

                units_list.innerHTML += div;
            });
            
        }else{

        }

        
    }

}
