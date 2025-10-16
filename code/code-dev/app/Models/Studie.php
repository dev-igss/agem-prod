<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class Studie extends Model
{
    use HasFactory;

    use SoftDeletes;

    protected $table = 'studies';
    protected $hidden = ['created_at', 'updated_at'];

    public function unit(){
        return $this->hasOne(Unit::class,'id','unit_id');
    }
}
