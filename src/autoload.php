<?php

spl_autoload_register(static function ($class) {
    $namespace = 'avadim\\FastExcelTemplator\\';
    if (0 === strpos($class, $namespace)) {
        include __DIR__ . '/FastExcelTemplator/' . str_replace($namespace, '', $class) . '.php';
    }
});

// EOF