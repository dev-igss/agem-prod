<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\SoftDeletes;

class Service extends Model
{
    use HasFactory;

    use SoftDeletes;

    protected $table = 'services';
    protected $hidden = ['created_at', 'updated_at'];

    public function parent(){
        return $this->hasOne(Service::class,'id','parent_id');
    }

    public function unit(){
        return $this->hasOne(Unit::class,'id','unit_id');
    }
}
