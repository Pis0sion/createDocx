<?php


namespace Pis0sion\Docx\entity;

use Pis0sion\Docx\service\MergeDocxService;
use RuntimeException;

/**
 * Class MergeDocx
 * @package Pis0sion\Docx\entity
 */
class MergeDocx
{
    /**
     * @var MergeDocxService
     */
    protected $docxService;

    /**
     * MergeDocx constructor.
     */
    public function __construct()
    {
        $this->docxService = new MergeDocxService();
    }

    /**
     * 合并文件
     * @param string $tempPath
     * @param string $file
     */
    public function run(string $tempPath, string $file)
    {
        $files = glob($tempPath . "/*.docx");
        $content = [];
        $r = '';
        for ($i = 1; $i < count($files); $i++) {
            // Open the all document - 1
            $this->docxService->Open($files[$i]);
            $content[$i] = $this->docxService->FileRead('word/document.xml');
            $this->docxService->Close();
            // Extract the content of  document
            $p = strpos($content[$i], '<w:body');
            if ($p === false) {
                throw new RuntimeException("Tag <w:body> not found in document ." . $files[$i]);
            }
            $p = strpos($content[$i], '>', $p);
            $content[$i] = substr($content[$i], $p + 1);
            $p = strpos($content[$i], '</w:body>');
            if ($p === false) {
                throw new RuntimeException("Tag <w:body> not found in document ." . $files[$i]);
            }
            //'<w:p><w:r><w:br w:type="page" /><w:lastRenderedPageBreak/></w:r></w:p>'.
            $content[$i] = substr($content[$i], 0, $p);
            $r .= $content[$i];
        }

        $this->docxService->Open($files[0]);
        $content2 = $this->docxService->FileRead('word/document.xml');
        $p = strpos($content2, '</w:body>');
        if ($p === false) {
            throw new RuntimeException("Tag <w:body> not found in document ." . $files[0]);
        }
        $content2 = substr_replace($content2, $r, $p, 0);
        $this->docxService->FileReplace('word/document.xml', $content2, TBSZIP_STRING);
        $this->docxService->Flush(TBSZIP_FILE, $file);
    }

    /**
     * 清除临时文件
     * @param string $dir_path
     */
    public function clearTemporaryFiles(string $dir_path)
    {
        if (is_dir($dir_path)) {
            $dirs = scandir($dir_path);
            foreach ($dirs as $dir) {
                if ($dir != '.' && $dir != '..') {
                    $sonDir = $dir_path . '/' . $dir;
                    if (is_dir($sonDir)) {
                        $this->clearTemporaryFiles($sonDir);
                        @rmdir($sonDir);
                    } else {
                        @unlink($sonDir);
                    }
                }
            }
            @rmdir($dir_path);
        }
    }

}