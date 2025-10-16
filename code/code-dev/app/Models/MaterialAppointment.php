<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class MaterialAppointment extends Model
{
    use HasFactory;

    protected $table = 'materials_appointments';
    protected $hidden = ['created_at', 'updated_at'];

    public function study(){
        return $this->hasOne(Studie::class,'id','idstudy');
    }
}
