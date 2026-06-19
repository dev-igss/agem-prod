 @extends('layouts.connect.master')

@section('title','Inicio de Sesión')

@section('content')
    <div class="box box_login shadow">

        <div class="header">
            <a href="{{url('/')}}">
                <img src="{{url('/img/logo igss.png')}}" alt="">
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

            @if ($errors->any())
                <div class="alert alert-danger mt-3">
                    @foreach ($errors->all() as $error)
                        <p class="mb-0">{{ $error }}</p>
                    @endforeach
                </div>
            @endif

            <div class="footer mtop16">
                <!-- <a href="{{url('/recover')}}">Recuperar Contraseña</a> -->
            </div>
        </div>

        
    </div>
@stop
