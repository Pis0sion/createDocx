<?php


namespace Pis0sion\Docx\component;

use PhpOffice\PhpWord\TemplateProcessor;

/**
 * Interface ICreateDocxTemplateInterface
 * @package Pis0sion\Docx\component
 */
interface ICreateDocxTemplateInterface
{
    /**
     * template processor
     * @param string $templatePath
     * @return TemplateProcessor
     */
    public function createDocxTemplate(string $templatePath): TemplateProcessor;

}