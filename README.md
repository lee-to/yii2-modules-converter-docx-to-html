# Модуль для Yii2 с помощью которого можно конвертировать docx файлы в html код

# Установка

1) composer require lee-to/yii2-modules-converter-docx-to-html
2) Скопировать wordhtml в папку с модулями yii2
3) Добавить в конфиг
'modules' => [
        'wordhtml' => [
            'class' => 'app\modules\wordhtml\Wordhtml',
        ],
	]
  
# License
The MIT License (MIT). Please see License File for more information.
