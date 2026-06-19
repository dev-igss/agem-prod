<?php

use Illuminate\Foundation\Application;
use Illuminate\Foundation\Configuration\Exceptions;
use Illuminate\Foundation\Configuration\Middleware;

return Application::configure(basePath: dirname(__DIR__))
    ->withRouting(
        web: __DIR__.'/../routes/web.php',
        commands: __DIR__.'/../routes/console.php',
        health: '/up',
        
    )
    ->withMiddleware(function (Middleware $middleware) {
        $middleware->alias([
            'Excel' => Maatwebsite\Excel\Facades\Excel::class,
            'user_perm' => \App\Http\Middleware\CheckUserPermission::class,
        ]);
    })
    ->withExceptions(function (Exceptions $exceptions) {
        $exceptions->render(function (\Illuminate\Session\TokenMismatchException $e, \Illuminate\Http\Request $request) {
            return redirect()->route('login')->withErrors(['ibm' => 'Tu sesión ha expirado. Por favor inicia sesión nuevamente.']);
        });

        $exceptions->render(function (\Illuminate\Auth\AuthenticationException $e, \Illuminate\Http\Request $request) {
            return redirect()->route('login');
        });
    })->create();
    