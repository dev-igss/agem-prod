<?php

namespace App\Http\Middleware;

use Closure;
use Illuminate\Http\Request;
use Symfony\Component\HttpFoundation\Response;

class CheckUserPermission
{
    /**
     * Handle an incoming request.
     *
     * @param  \Closure(\Illuminate\Http\Request): (\Symfony\Component\HttpFoundation\Response)  $next
     */
    public function handle(Request $request, Closure $next): Response
    {
        if (!Auth::check()) {
            return redirect()->route('login')->with('error', 'Tu sesión ha expirado por inactividad.');
        }

        // 2. Validar si el usuario tiene el permiso usando tu función personalizada
        // Asumiendo que kvfj recibe (datos_json, clave_a_buscar)
        if (!kvfj(Auth::user()->permissions)) {
            return redirect()->route('admin.dashboard')->with('error', 'No tienes permiso para acceder a esta sección.');
        }
    }
}
