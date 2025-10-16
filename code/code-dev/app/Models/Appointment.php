<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class Appointment extends Model
{
    use HasFactory;

    use SoftDeletes;

    protected $table = 'appointments';
    protected $hidden = ['created_at', 'updated_at'];

    public function patient(){
        return $this->hasOne(Patient::class,'id','patient_id');
    }

    public function service(){
        return $this->hasOne(Service::class,'id','service_id');
    }

    public function schedule(){
        return $this->hasOne(Schedule::class,'id','schedule_id');
    }

    public function studie(){
        return $this->hasOne(Studie::class,'id','study_id');
    }

    public function details(){
        return $this->hasMany(DetailAppointment::class,'idappointment', 'id');
    }

    public function materials(){
        return $this->hasOne(MaterialAppointment::class,'idappointment', 'id');
    }

    public function tecnico1(){
        return $this->hasOne(User::class,'id','ibm_tecnico_1');
    }

    public function tecnico2(){
        return $this->hasOne(User::class,'id','ibm_tecnico_2');
    }
}
