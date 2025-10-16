<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class CodePatient extends Model
{
    use HasFactory;
    use SoftDeletes;

    protected $table = 'codes_patients';
    protected $hidden = ['created_at', 'updated_at'];

    public function patient(){
        return $this->hasOne(Patient::class,'id','patient_id');
    }
}
