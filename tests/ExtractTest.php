<?php

namespace Ottosmops\Office2text\Test;

use PHPUnit\Framework\TestCase;
use Ottosmops\Office2text\Extract;
use Ottosmops\Office2text\Exceptions\FileNotFound;
use Ottosmops\Office2text\Exceptions\CouldNotExtractText;
use Ottosmops\Office2text\Exceptions\ExtractException;

class ExtractTest extends TestCase
{
    public function testCanExtractTextFromDocx(): void
    {
        // Create a minimal valid docx zip file for testing
        $testFile = __DIR__ . '/files/minimal.docx';
        $this->createMinimalDocx($testFile);
        
        $text = (new Extract())
            ->document($testFile)
            ->text();
        
        $this->assertStringContainsString('Test Document Content', $text);
        
        unlink($testFile); // cleanup
    }

    public function testProvidesStaticMethodToExtractText(): void
    {
        $testFile = __DIR__ . '/files/static.docx';
        $this->createMinimalDocx($testFile);
        
        $text = Extract::getText($testFile);
        $this->assertStringContainsString('Test Document Content', $text);
        
        unlink($testFile);
    }

    public function testCanExtractTextFromPptx(): void
    {
        $testFile = __DIR__ . '/files/minimal.pptx';
        $this->createMinimalPptx($testFile);
        
        $text = (new Extract())
            ->document($testFile)
            ->text();
        
        $this->assertStringContainsString('Test Slide Content', $text);
        
        unlink($testFile);
    }

    public function testCanExtractTextFromXlsx(): void
    {
        $testFile = __DIR__ . '/files/minimal.xlsx';
        $this->createMinimalXlsx($testFile);
        
        $text = (new Extract())
            ->document($testFile)
            ->text();
        
        $this->assertStringContainsString('Test Cell Content', $text);
        
        unlink($testFile);
    }

    public function testThrowsExceptionWhenFileIsNotFound(): void
    {
        $this->expectException(FileNotFound::class);
        (new Extract())
            ->document('/no/document/here/dummy.docx')
            ->text();
    }

    public function testThrowsExceptionWhenFileTypeIsNotSupported(): void
    {
        $this->expectException(CouldNotExtractText::class);
        (new Extract())
            ->document(__DIR__ . '/files/test.txt')
            ->text();
    }

    public function testCanExtractTextFromOdt(): void
    {
        $testFile = __DIR__ . '/files/minimal.odt';
        $this->createMinimalOdt($testFile);
        
        $text = (new Extract())
            ->document($testFile)
            ->text();
        
        $this->assertStringContainsString('Test ODT Content', $text);
        
        unlink($testFile);
    }

    public function testCanExtractTextFromOdp(): void
    {
        $testFile = __DIR__ . '/files/minimal.odp';
        $this->createMinimalOdp($testFile);
        
        $text = (new Extract())
            ->document($testFile)
            ->text();
        
        $this->assertStringContainsString('Test ODP Content', $text);
        
        unlink($testFile);
    }

    public function testCanExtractTextFromOds(): void
    {
        $testFile = __DIR__ . '/files/minimal.ods';
        $this->createMinimalOds($testFile);
        
        $text = (new Extract())
            ->document($testFile)
            ->text();
        
        $this->assertStringContainsString('Test ODS Content', $text);
        
        unlink($testFile);
    }
    
    private function createMinimalDocx($filename)
    {
        $zip = new \ZipArchive();
        $zip->open($filename, \ZipArchive::CREATE);
        
        // Add minimal document.xml
        $documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Test Document Content</w:t></w:r></w:p>
    </w:body>
</w:document>';
        
        $zip->addFromString('word/document.xml', $documentXml);
        
        // Add minimal _rels/.rels
        $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>';
        
        $zip->addFromString('_rels/.rels', $relsXml);
        
        // Add Content_Types
        $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>';
        
        $zip->addFromString('[Content_Types].xml', $contentTypes);
        $zip->close();
    }

    private function createMinimalPptx($filename)
    {
        $zip = new \ZipArchive();
        $zip->open($filename, \ZipArchive::CREATE);
        
        // Add minimal slide1.xml
        $slideXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:sp>
                <p:txBody>
                    <a:p>
                        <a:r>
                            <a:t>Test Slide Content</a:t>
                        </a:r>
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:sld>';
        
        $zip->addFromString('ppt/slides/slide1.xml', $slideXml);
        
        // Add minimal _rels/.rels
        $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>';
        
        $zip->addFromString('_rels/.rels', $relsXml);
        
        // Add Content_Types
        $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>';
        
        $zip->addFromString('[Content_Types].xml', $contentTypes);
        $zip->close();
    }

    private function createMinimalXlsx($filename)
    {
        $zip = new \ZipArchive();
        $zip->open($filename, \ZipArchive::CREATE);
        
        // Add minimal sheet1.xml
        $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr">
                <is>
                    <t>Test Cell Content</t>
                </is>
            </c>
        </row>
    </sheetData>
</worksheet>';
        
        $zip->addFromString('xl/worksheets/sheet1.xml', $sheetXml);
        
        // Add minimal _rels/.rels
        $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>';
        
        $zip->addFromString('_rels/.rels', $relsXml);
        
        // Add Content_Types
        $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>';
        
        $zip->addFromString('[Content_Types].xml', $contentTypes);
        $zip->close();
    }

    private function createMinimalOdt($filename)
    {
        $zip = new \ZipArchive();
        $zip->open($filename, \ZipArchive::CREATE);
        
        // Add minimal content.xml for ODT
        $contentXml = '<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:text>
            <text:p>Test ODT Content</text:p>
        </office:text>
    </office:body>
</office:document-content>';
        
        $zip->addFromString('content.xml', $contentXml);
        
        // Add minimal META-INF/manifest.xml
        $manifestXml = '<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
    <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>';
        
        $zip->addFromString('META-INF/manifest.xml', $manifestXml);
        $zip->close();
    }

    private function createMinimalOdp($filename)
    {
        $zip = new \ZipArchive();
        $zip->open($filename, \ZipArchive::CREATE);
        
        // Add minimal content.xml for ODP
        $contentXml = '<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
    <office:body>
        <office:presentation>
            <draw:page>
                <draw:frame>
                    <draw:text-box>
                        <text:p>Test ODP Content</text:p>
                    </draw:text-box>
                </draw:frame>
            </draw:page>
        </office:presentation>
    </office:body>
</office:document-content>';
        
        $zip->addFromString('content.xml', $contentXml);
        
        // Add minimal META-INF/manifest.xml
        $manifestXml = '<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>
    <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>';
        
        $zip->addFromString('META-INF/manifest.xml', $manifestXml);
        $zip->close();
    }

    private function createMinimalOds($filename)
    {
        $zip = new \ZipArchive();
        $zip->open($filename, \ZipArchive::CREATE);
        
        // Add minimal content.xml for ODS
        $contentXml = '<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
    <office:body>
        <office:spreadsheet>
            <table:table>
                <table:table-row>
                    <table:table-cell>
                        <text:p>Test ODS Content</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table>
        </office:spreadsheet>
    </office:body>
</office:document-content>';
        
        $zip->addFromString('content.xml', $contentXml);
        
        // Add minimal META-INF/manifest.xml
        $manifestXml = '<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
    <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>';
        
        $zip->addFromString('META-INF/manifest.xml', $manifestXml);
        $zip->close();
    }
}
