<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Bitacora extends Model
{
    use HasFactory;

    protected $table = 'bitacoras';
    protected $hidden = ['created_at', 'updated_at'];

    public function user(){
        return $this->hasOne(User::class,'id','user_id');
    }

    public function ing7(){
        return $this->hasOne(Ings7::class,'id','iding7');
    }
}
