<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class Patient extends Model
{
    use HasFactory;

    use SoftDeletes;

    protected $table = 'patients';
    protected $hidden = ['created_at', 'updated_at'];
    protected $fillable = [
        'id',
        'unit_id',
        'exp_prev',
        'type',
        'affiliation',
        'affiliation_principal',
        'affiliation_idparent',
        'name',
        'lastname',
        'contact',
        'birth',
        'age',
        'gender',
        'deleted_at',
        'created_at',
        'updated_at',
    ];

    public function unit(){
        return $this->hasOne(Unit::class,'id','unit_id');
    }

    public function parent(){
        return $this->hasOne(Patient::class,'id','affiliation_idparent');
    }
}
