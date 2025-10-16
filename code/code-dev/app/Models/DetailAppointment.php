<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class DetailAppointment extends Model
{
    use HasFactory;

    protected $table = 'details_appointments';
    protected $hidden = ['created_at', 'updated_at'];

    public function service(){
        return $this->hasOne(Service::class,'id','idservice');
    }

    public function study(){
        return $this->hasOne(Studie::class,'id','idstudy');
    }

    public function materials(){
        return $this->hasOne(MaterialAppointment::class,'idappointment','idappointment');
    }
}
