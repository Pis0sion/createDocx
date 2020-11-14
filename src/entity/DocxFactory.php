<?php


namespace Pis0sion\Docx\entity;

use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\TemplateProcessor;
use Pis0sion\Docx\service\CreateDocxTemplateService;
use Pis0sion\Docx\service\DocxRenderTemplateService;

/**
 * Class DocxFactory
 * @package Pis0sion\Docx\entity
 */
class DocxFactory
{
    /**
     * @var TemplateProcessor
     */
    protected $templateProcessor;

    /**
     * @var int
     */
    protected $isCover = 1;

    /**
     * @var int
     */
    protected $tocBlock = 1;

    /**
     * @var int
     */
    protected $brLine = '<w:p><w:r><w:br w:type="page" /><w:lastRenderedPageBreak/></w:r></w:p>';

    /**
     * @var int
     */
    protected $apiVersion = 1;

    /**
     * @var int
     */
    protected $apiList = 1;

    /**
     * @param array $renderData 渲染的数据
     * @param string $templatePath 模板的地址
     * @param string $saveAsPath 生成的文件地址
     * @throws CreateTemporaryFileException
     */
    public function run(array $renderData, string $templatePath, string $saveAsPath)
    {
        // 临时文件
        $temporaryPath = "./tmp" . DIRECTORY_SEPARATOR . uniqid();

        foreach ($renderData as $key => $datum) {
            // 设置模板
            $this->setTemplateProcessor($templatePath);
            // 初始化模板
            $this->initTemplateProcessor($key);
            // 渲染模板
            $this->renderTemplate($temporaryPath, $datum);
        }

        $this->mergeDocxAndClearTemporaryFiles($temporaryPath, $saveAsPath);

        echo "操作完成";
    }

    /**
     * @return TemplateProcessor
     */
    public function getTemplateProcessor(): TemplateProcessor
    {
        return $this->templateProcessor;
    }

    /**
     * @param string $templatePath
     * @throws CreateTemporaryFileException
     */
    public function setTemplateProcessor(string $templatePath): void
    {
        $this->templateProcessor = (new CreateDocxTemplateService())->createDocxTemplate($templatePath);
    }

    /**
     * 初始化模板
     * @param int $init
     */
    protected function initTemplateProcessor(int $init)
    {
        if ($init >= 1) {
            $this->isCover = 0;
            $this->tocBlock = 0;
            $this->brLine = '';
            $this->apiVersion = 0;
            $this->apiList = 0;
        }

        $this->templateProcessor->cloneBlock("Cover", $this->isCover);
        $this->templateProcessor->cloneBlock("TocBlock", $this->tocBlock);
        $this->templateProcessor->setValue("BrLine", $this->brLine);
        $this->templateProcessor->cloneBlock("ApiVersion", $this->apiVersion);
        $this->templateProcessor->cloneBlock("ApiList", $this->apiList);
    }

    /**
     * 渲染数据
     * @param $tmpPath
     * @param $datum
     */
    protected function renderTemplate($tmpPath, $datum)
    {
        // 渲染服务
        $docxService = new DocxRenderTemplateService($this->templateProcessor);
        // 保存到临时文件夹中
        $docxService->setSavePath($tmpPath);
        // 渲染数据
        $docxService->renderDataToTemplate($datum);
    }

    /**
     * 合并并删除临时文件
     * @param $temporaryPath
     * @param $saveAsPath
     */
    protected function mergeDocxAndClearTemporaryFiles($temporaryPath, $saveAsPath)
    {
        $mergeService = new MergeDocx();
        // 合并 docx
        $mergeService->run($temporaryPath, $saveAsPath);
        // 清空临时文件
        $mergeService->clearTemporaryFiles($temporaryPath);
    }
}