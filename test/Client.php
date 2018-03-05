<?php
/**
 * project: iningmeng.
 * Author: iningmeng
 * Email: iningmeng@qq.com
 * Date: 2018/3/5  14:12
 * Version: 1.0
 * Description:
 */
namespace iningmeng\pptimages;
    require __DIR__ . '/../vendor/autoload.php';
//    use iningmeng\pptimages\Handle;

    $handle = new Handle();
    $handle->index('../test/ppt/1.ppt','../test/images');