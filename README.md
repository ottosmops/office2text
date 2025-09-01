# office2text

[![CI](https://github.com/ottosmops/office2text/actions/workflows/ci.yml/badge.svg)](https://github.com/ottosmops/office2text/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/ottosmops/office2text/branch/main/graph/badge.svg)](https://codecov.io/gh/ottosmops/office2text)
[![Software License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)](LICENSE.md)
[![Latest Stable Version](https://poser.pugx.org/ottosmops/office2text/v/stable?format=flat-square)](https://packagist.org/packages/ottosmops/office2text)
[![Packagist Downloads](https://img.shields.io/packagist/dt/ottosmops/office2text.svg?style=flat-square)](https://packagist.org/packages/ottosmops/office2text)

Extract text from Microsoft Office (docx, pptx, xlsx) and LibreOffice (odt, odp, ods) documents using pure PHP (ZipArchive + SimpleXML). 

## Installation

With Composer:

```bash
composer require ottosmops/office2text
```

## Usage

```php
use Ottosmops\Office2text\Extract;

$text = (new Extract())
  ->document('example.docx')
  ->text();
```

Or using the static method:

```php
$text = Extract::getText('example.docx');
```

Supported formats: 
- **Microsoft Office**: docx, pptx, xlsx
- **LibreOffice**: odt, odp, ods

## License

[MIT License](LICENSE.md)
