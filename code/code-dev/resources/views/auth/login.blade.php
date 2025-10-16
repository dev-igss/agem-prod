 @extends('layouts.connect.master')

@section('title','Inicio de Sesión')

@section('content')
    <div class="box box_login shadow">

        <div class="header">
            <a href="{{url('/')}}">
                <img src="{{url('/static/imagenes/igss.png')}}" alt="">
            </a>
        </div>

        <div class="inside">
            <form method="POST" action="{{ route('login') }}">
                @csrf

                <label for="ibm"><strong>IBM:</strong></label>
                <div class="input-group">
                    <div class="input-group-prepend">
                    <div class="input-group-text"><i class="fas fa-address-card"></i></div>
                    </div>
                    <input type="text" class="form-control" name="ibm">
                </div>

                <label for="password" class="mtop16"><strong>Contraseña:</strong></label>
                <div class="input-group">
                    <div class="input-group-prepend">
                    <div class="input-group-text"><i class="fas fa-lock"></i></div>
                    </div>
                    <input type="password" class="form-control" name="password">
                </div>

                <button class="btn btn-success mt-4" type="submit">Inicar Sesión</button>
            </form>

            @if(Session::has('message'))
                <div class="container">
                    <div class="mt-16 alert alert-{{ Session::get('typealert') }}" style="display:none;">
                        {{ Session::get('message') }}
                        @if( $errors->any() )
                            <ul>
                                @foreach($errors->all() as $error)
                                    <li>{{ $error }}</li>
                                @endforeach
                            </ul>
                        @endif
                        <script>
                            $('.alert').slideDown();
                            setTimeout(function(){ $('.alert').slideUp(); },10000);
                        </script>
                    </div>
                </div>
            @endif

            <div class="footer mtop16">
                <!-- <a href="{{url('/recover')}}">Recuperar Contraseña</a> -->
            </div>
        </div>

        
    </div>
@stop