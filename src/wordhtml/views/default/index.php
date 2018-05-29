<?
use yii\widgets\ActiveForm;
use yii\helpers\Html;
?>

<div class="page-header">
    <h3>Конвертер DOCX в HTML</h3>
</div>

<div class="row">
    <div class="col-md-10 col-sm-10">
        <?php $form = ActiveForm::begin(['options' => ['name' => 'docForm', 'enctype' => 'multipart/form-data']]) ?>
            <div class="form-group">
                <?= $form->field($model, 'file')->fileInput() ?>
                <p class="help-block">Только файлы формата docx</p>
            </div>
            <?= $form->field($model, 'save_type_html')->checkbox() ?>
            <?= $form->field($model, 'save_type_xls')->checkbox() ?>

            <?= Html::submitButton('Конвертировать', ['class' => 'btn btn-default']) ?>
        <?php ActiveForm::end() ?>
    </div>
</div>


<? if($textareaView):?>
    <hr>
    <h3>Результат</h3>
    <h5><?=mb_strlen(strip_tags($textareaView),'utf-8');?> символ(ов)</h5>
    <h5><?=str_word_count(strip_tags($textareaView));?> слов(а)</h5>
    <hr>
    <h4>Код</h4>
    <hr>
    <textarea id="htmlviewer" class="form-control" style="width: 100%;margin-top: 20px;" rows="20" cols="2"><?=$textareaView;?></textarea>
    <h4>Визуальный просмотр</h4>
    <hr>
    <?=$textareaView;?>
    <hr>
<?endif;?>