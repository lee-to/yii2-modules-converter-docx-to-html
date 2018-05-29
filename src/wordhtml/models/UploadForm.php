<?
namespace app\modules\wordhtml\models;

use yii\base\Model;
use yii\web\UploadedFile;

/**
 * UploadForm is the model behind the upload form.
 */
class UploadForm extends Model
{
    /**
     * @var UploadedFile file attribute
     */
    public $file;
    public $save_type_html;
    public $save_type_xls;

    /**
     * @return array the validation rules.
     */
    public function rules()
    {
        return [
            [['file'], 'file', 'skipOnEmpty' => false, 'extensions' => 'docx'],
            [['save_type_html', 'save_type_xls'], 'string']
        ];
    }

    public function attributeLabels()
    {
        return [
            'file' => 'Файл Docx',
            'save_type_html' => 'Сохранить результат как html файл',
            'save_type_xls' => 'Сохранить результат как excel файл',
        ];
    }
}