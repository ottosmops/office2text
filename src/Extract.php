<?php

namespace Ottosmops\Office2text;

use Ottosmops\Office2text\Exceptions\ExtractException;
use Ottosmops\Office2text\Exceptions\FileNotFound;
use Ottosmops\Office2text\Exceptions\CouldNotExtractText;

class Extract
{
    /**
     * @var string Path to the source file
     */
    protected $source = '';

    /**
     * @var string Extracted text content
     */
    private $text = '';

    /**
     * Setup
     */
    public function __construct()
    {
        // Empty constructor like pdftotext
    }

    /**
     * Get text from document
     * @param  string $source
     * @return string
     */
    public static function getText($source)
    {
        return (new static())
                  ->document($source)
                  ->text();
    }

    /**
     * Set document file (source)
     * @param  string $source
     * @return $this
     * @throws FileNotFound
     */
    public function document($source)
    {
        if (!file_exists($source)) {
            throw new FileNotFound("Could not find document file: {$source}");
        }
        $this->source = $source;
        return $this;
    }

    /**
     * Extract text
     * @return string
     * @throws CouldNotExtractText
     */
    public function text()
    {
        try {
            return $this->extractText();
        } catch (\Exception $e) {
            throw new CouldNotExtractText($e->getMessage(), 0, $e);
        }
    }

    private function extractText()
    {
        $ext = strtolower(pathinfo($this->source, PATHINFO_EXTENSION));
        switch ($ext) {
            case 'docx':
                return $this->extractDocx();
            case 'pptx':
                return $this->extractPptx();
            case 'xlsx':
                return $this->extractXlsx();
            case 'odt':
                return $this->extractOdt();
            case 'odp':
                return $this->extractOdp();
            case 'ods':
                return $this->extractOds();
            default:
                throw new ExtractException("Unsupported file type: $ext");
        }
    }

    private function extractDocx()
    {
        $zip = new \ZipArchive();
        if ($zip->open($this->source) === TRUE) {
            $xml = $zip->getFromName('word/document.xml');
            $zip->close();
            if ($xml === false) throw new ExtractException('word/document.xml not found');
            $xml = preg_replace('/<w:tab\/?>(.*?)<\/w:tab>/', "\t", $xml); // Tabs
            $xml = preg_replace('/<w:br\/?>(.*?)<\/w:br>/', "\n", $xml); // Line breaks
            $xml = str_replace(['<w:p>', '</w:p>'], ["", "\n"], $xml); // Paragraphs
            $text = strip_tags($xml);
            return html_entity_decode($text, ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        throw new ExtractException('Could not open file as zip');
    }

    private function extractPptx()
    {
        $zip = new \ZipArchive();
        $text = '';
        if ($zip->open($this->source) === TRUE) {
            // Iterate over all slides
            for ($i = 1; ; $i++) {
                $slideName = sprintf('ppt/slides/slide%d.xml', $i);
                $xml = $zip->getFromName($slideName);
                if ($xml === false) break;
                $sxml = simplexml_load_string($xml);
                if ($sxml === false) continue;
                $texts = $sxml->xpath('//a:t');
                foreach ($texts as $t) {
                    $text .= (string)$t . "\n";
                }
            }
            $zip->close();
            return $text;
        }
        throw new ExtractException('Could not open file as zip');
    }

    private function extractXlsx()
    {
        $zip = new \ZipArchive();
        $text = '';
        if ($zip->open($this->source) === TRUE) {
            $sharedStrings = [];
            $sharedXml = $zip->getFromName('xl/sharedStrings.xml');
            if ($sharedXml !== false) {
                $sxml = simplexml_load_string($sharedXml);
                foreach ($sxml->si as $si) {
                    $sharedStrings[] = (string)$si->t;
                }
            }
            // Iterate over all worksheets
            for ($i = 1; ; $i++) {
                $sheetName = sprintf('xl/worksheets/sheet%d.xml', $i);
                $xml = $zip->getFromName($sheetName);
                if ($xml === false) break;
                $sxml = simplexml_load_string($xml);
                if ($sxml === false) continue;
                
                foreach ($sxml->sheetData->row as $row) {
                    $rowText = [];
                    foreach ($row->c as $c) {
                        $v = '';
                        // Check for inline strings
                        if (isset($c->is->t)) {
                            $v = (string)$c->is->t;
                        }
                        // Check for shared strings
                        elseif (isset($c['t']) && $c['t'] == 's') {
                            $index = (int)$c->v;
                            $v = $sharedStrings[$index] ?? '';
                        }
                        // Regular cell value
                        elseif (isset($c->v)) {
                            $v = (string)$c->v;
                        }
                        
                        if ($v !== '') {
                            $rowText[] = $v;
                        }
                    }
                    if (!empty($rowText)) {
                        $text .= implode("\t", $rowText) . "\n";
                    }
                }
            }
            $zip->close();
            return $text;
        }
        throw new ExtractException('Could not open file as zip');
    }

    private function extractOdt()
    {
        $zip = new \ZipArchive();
        if ($zip->open($this->source) === TRUE) {
            $xml = $zip->getFromName('content.xml');
            $zip->close();
            if ($xml === false) throw new ExtractException('content.xml not found');
            
            // Remove LibreOffice formatting tags and extract text
            $xml = preg_replace('/<text:tab\/>/', "\t", $xml); // Tabs
            $xml = preg_replace('/<text:line-break\/>/', "\n", $xml); // Line breaks
            $xml = str_replace(['<text:p', '</text:p>'], ["<text:p", "\n"], $xml); // Paragraphs
            $text = strip_tags($xml);
            return html_entity_decode($text, ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        throw new ExtractException('Could not open file as zip');
    }

    private function extractOdp()
    {
        $zip = new \ZipArchive();
        $text = '';
        if ($zip->open($this->source) === TRUE) {
            $xml = $zip->getFromName('content.xml');
            $zip->close();
            if ($xml === false) throw new ExtractException('content.xml not found');
            
            $sxml = simplexml_load_string($xml);
            if ($sxml === false) throw new ExtractException('Invalid XML content');
            
            // Register namespaces for LibreOffice
            $sxml->registerXPathNamespace('text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
            $sxml->registerXPathNamespace('draw', 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0');
            
            // Extract text from all text elements
            $texts = $sxml->xpath('//text:p | //text:span | //text:h');
            foreach ($texts as $t) {
                $text .= (string)$t . "\n";
            }
            return $text;
        }
        throw new ExtractException('Could not open file as zip');
    }

    private function extractOds()
    {
        $zip = new \ZipArchive();
        $text = '';
        if ($zip->open($this->source) === TRUE) {
            $xml = $zip->getFromName('content.xml');
            $zip->close();
            if ($xml === false) throw new ExtractException('content.xml not found');
            
            $sxml = simplexml_load_string($xml);
            if ($sxml === false) throw new ExtractException('Invalid XML content');
            
            // Register namespaces for LibreOffice Calc
            $sxml->registerXPathNamespace('table', 'urn:oasis:names:tc:opendocument:xmlns:table:1.0');
            $sxml->registerXPathNamespace('text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0');
            
            // Extract text from all table cells
            $cells = $sxml->xpath('//table:table-cell//text:p');
            $rowData = [];
            
            foreach ($cells as $cell) {
                $cellText = trim((string)$cell);
                if ($cellText !== '') {
                    $rowData[] = $cellText;
                }
            }
            
            if (!empty($rowData)) {
                $text = implode("\t", $rowData) . "\n";
            }
            
            return $text;
        }
        throw new ExtractException('Could not open file as zip');
    }
}
