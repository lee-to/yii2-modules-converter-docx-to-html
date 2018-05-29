<?php

namespace app\modules\wordhtml\controllers;

use yii\web\Controller;
use app\models\User;
use yii\web\ForbiddenHttpException;
use wordhtml\components\WordParser;
use app\modules\wordhtml\models\UploadForm;
use yii\web\UploadedFile;

class DefaultController extends \app\components\Controller
{
    public function actionIndex()
    {
        $WordParser = new WordParser();
        $model = new UploadForm();
        $post = \Yii::$app->request->post();

        $textareaView = '';

        if (\Yii::$app->request->isPost) {
            $model->file = UploadedFile::getInstance($model, 'file');

            if ($model->file && $model->validate()) {
                $filename = uniqid('file_').'.'.$model->file->extension;
                if(!$model->file->saveAs('web/uploads/wordhtml/'.$filename)){
                    \Yii::$app->getSession()->setFlash('error', 'Файл не загружен.');
                }
                else {
                    $original = 'web/uploads/wordhtml/'.$filename;

                    $WordParser->file = $original;
                    $WordParser->tmpDir = 'tmp';
                    $WordParser->init();

                    $generate_urls = false;

                    $model->load($post);
                    if($model->save_type_html == 1) {
                        $WordParser->sendDownloadFile(true, $generate_urls);
                    }
                    elseif($model->save_type_xls == 1){
                        $WordParser->generateXls($generate_urls);
                    }
                    else {
                        $textareaView = $WordParser->getHtml(false, $generate_urls);
                    }

                    unlink($original);
                }
            }
            else {
                \Yii::$app->getSession()->setFlash('error', 'Ошибка валидации.');
            }
        }

        return $this->render('index', [
            'model' => $model,
            'textareaView' => $textareaView
        ]);
    }
}
