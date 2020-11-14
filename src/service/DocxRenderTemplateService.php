<?php


namespace Pis0sion\Docx\service;

use Exception;
use PhpOffice\PhpWord\TemplateProcessor;
use RuntimeException;
use SplQueue;

const Default_Description = "暂无描述";

/**
 * Class DocxRenderTemplateService
 * @package Pis0sion\Docx\service
 */
class DocxRenderTemplateService
{
    /**
     * @var TemplateProcessor
     */
    protected $docxTemplate;

    /**
     * @var SplQueue
     */
    protected $procedure;

    /**
     * @var string
     */
    protected $savePath;

    /**
     * @var string
     */
    protected $ext;

    /**
     * @var array
     */
    protected $ReqParameters = [
        "HeaderVars", "QueryVars", "BodyVars", "RawVars",
    ];

    /**
     * @var array
     */
    protected $RespParameters = [
        "RespBody", "RespRawVars",
    ];

    /**
     * DocxRenderTemplateService constructor.
     * @param TemplateProcessor $docxTemplate
     */
    public function __construct(TemplateProcessor $docxTemplate)
    {
        $this->docxTemplate = $docxTemplate;
        $this->ext = ".docx";
    }

    /**
     * 渲染数据到模板
     * @param array $renderData
     */
    public function renderDataToTemplate(array $renderData)
    {
        $this->docxTemplate = $this->pipeline($renderData);
        $this->docxTemplate->saveAs($this->savePath . DIRECTORY_SEPARATOR . uniqid() . $this->ext);
    }

    /**
     * 设置保存生成 docx 的文件
     * @param string $filePath
     */
    public function setSavePath(string $filePath)
    {
        if (!is_dir($filePath)) {
            @mkdir($filePath);
        }
        $this->savePath = $filePath;
    }

    /**
     * 获取保存的 docx 的路径
     * @return string
     */
    public function getSavePath(): string
    {
        return $this->savePath;
    }

    /**
     * @param array $renderData
     * @return TemplateProcessor
     */
    protected function pipeline(array $renderData)
    {
        $this->prepareToPipeline(
            [$this, "renderDocumentInformation"],
            [$this, "renderApiBaseInfo"],
            [$this, "renderRequestParameters"],
            [$this, "renderResponseParameters"],
        );

        while (!$this->procedure->isEmpty()) {
            $this->docxTemplate = $this->procedure->dequeue()($renderData);
        }

        return $this->docxTemplate;
    }

    /**
     * @param callable ...$callables
     * @return void
     */
    protected function prepareToPipeline(callable ...$callables)
    {
        $this->procedure = new SplQueue();
        foreach ($callables as $callable) {
            $this->procedure->enqueue($callable);
        }
    }

    /**
     * @param $parameter
     * @return TemplateProcessor
     */
    protected function renderDocumentInformation($parameter)
    {
        $documentVars['Title'] = $parameter['Title'];
        $documentVars['UpdateTime'] = $parameter['UpdateTime'];
        $documentVars['Description'] = $this->achieveDescription($parameter["Description"]);

        foreach ($documentVars as $key => $documentVar) {
            $this->docxTemplate->setValue($key, $documentVar);
        }

        return $this->docxTemplate;
    }

    /**
     * 渲染接口基本信息
     * @param $parameter
     * @return TemplateProcessor
     */
    protected function renderApiBaseInfo($parameter)
    {
        if (!array_key_exists('ApiBaseInfo', $parameter)) {
            throw new RuntimeException("ApiBaseInfo key isn't exist");
        }

        foreach ($parameter['ApiBaseInfo'] as $key => $value) {
            $this->docxTemplate->setValue($key, $value);
        }

        return $this->docxTemplate;
    }

    /**
     * 渲染请求参数
     * @param $parameter
     * @return TemplateProcessor
     * @throws Exception
     */
    protected function renderRequestParameters(array $parameter)
    {
        return $this->renderParameters($parameter, "ReqParameters");
    }

    /**
     * 渲染响应参数
     * @param $parameter
     * @return TemplateProcessor
     * @throws Exception
     */
    protected function renderResponseParameters(array $parameter)
    {
        return $this->renderParameters($parameter, "RespParameters");
    }

    /**
     * 渲染参数
     * @param array $parameter
     * @param string $module
     * @return TemplateProcessor
     * @throws Exception
     */
    protected function renderParameters(array $parameter, string $module)
    {
        if (!array_key_exists('ReqParameters', $parameter)) {
            throw new RuntimeException("ReqParameters key isn't exist");
        }

        $renderKeys = array_keys($parameter[$module]);

        // 需要删除的block
        $removeBlocks = array_diff($this->{$module}, $renderKeys);

        foreach ($removeBlocks as $removeBlock) {
            $this->docxTemplate->cloneBlock($removeBlock, 0);
        }

        //  渲染自己的表格
        foreach ($renderKeys as $renderKey) {
            //  渲染简单的基础数据
            if (!is_array($parameter[$module][$renderKey])) {
                $this->handleSimpleValues($renderKey, $parameter[$module][$renderKey]);
                continue;
            }
            //  渲染复杂的数据格式
            $this->handleComplexValues($renderKey, $parameter[$module][$renderKey]);
        }

        return $this->docxTemplate;
    }

    /**
     * 处理数组的简单逻辑
     * @param string $block
     * @param $value
     * @return TemplateProcessor
     */
    protected function handleSimpleValues(string $block, $value)
    {
        if (!empty($value)) {
            $this->docxTemplate->cloneBlock($block);
            $rKey = rtrim($block, "Vars");
            $this->docxTemplate->setValue($rKey, $this->prettify($value));
        } else {
            $this->docxTemplate->cloneBlock($block, 0);
        }

        return $this->docxTemplate;
    }

    /**
     * 处理数组的复杂逻辑
     * @param string $block
     * @param array $values
     * @return TemplateProcessor
     * @throws Exception
     */
    protected function handleComplexValues(string $block, array $values)
    {
        $multiLines = count($values);

        $this->docxTemplate->cloneBlock($block);
        $this->docxTemplate->cloneRow(key(current($values)), $multiLines);
        for ($j = 0; $j < $multiLines; $j++) {
            foreach ($values[$j] as $key => $value) {
                $this->docxTemplate->setValue($key . "#" . ($j + 1), htmlentities($value));
            }
        }
        return $this->docxTemplate;
    }

    /**
     * 获取描述
     * @param string $description
     * @return string
     */
    protected function achieveDescription(string $description): string
    {
        if (is_string($description) && trim($description) == "") {
            $description = Default_Description;
        }
        if (is_bool($description)) {
            $description = Default_Description;
        }
        return $description;
    }

    /**
     * @param string|null $regex
     * @return string
     */
    protected function prettify(?string $regex): string
    {
        // 验证是否为json数据
        $originStr = json_decode($regex, true);
        // 判断是否json序列化的成功
        if (json_last_error()) {
            return htmlentities($regex);
        }
        // json 数据做美化
        $regex = json_encode($originStr, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
        return str_replace(array("\r\n", "\r", "\n"), "<w:br />" . "\r", htmlentities($regex));
    }
}