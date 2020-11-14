<?php


namespace Pis0sion\Docx\service;


use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\TemplateProcessor;
use Pis0sion\Docx\component\ICreateDocxTemplateInterface;

/**
 * Class CreateDocxTemplateService
 * @package Pis0sion\Docx\service
 */
class CreateDocxTemplateService implements ICreateDocxTemplateInterface
{

    /**
     * @param string $templatePath
     * @return TemplateProcessor
     * @throws CreateTemporaryFileException
     */
    public function createDocxTemplate(string $templatePath): TemplateProcessor
    {
        // TODO: Implement createDocxTemplate() method.
        try {
            return new TemplateProcessor($templatePath);
        } catch (\Throwable $throwable) {
            throw new CreateTemporaryFileException();
        }
    }
}