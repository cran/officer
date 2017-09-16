officer R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->
The officer package lets R users manipulate Word (`.docx`) and PowerPoint (`*.pptx`) documents. In short, one can add images, tables and text into documents from R. An initial document can be provided, contents, styles and properties of the original document will then be available.

*This package is close to ReporteRs as it produces Word and PowerPoint files but it is faster, do not require `rJava` (but `xml2`) and has less functions that will make it easier to maintain.*

<!--html_preserve-->
<span><img src="data:image/png;base64,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"/></span><!--/html_preserve-->

> Make corporate reporting with minimum hassle

Word documents
--------------

Function `read_docx` will read an initial Word document (an empty one by default) and let you modify its content later.

The package provides functions to add R outputs into a Word document:

-   images: produce your plot in png or emf files and add them into the document; as a whole paragraph or inside a paragraph.
-   tables: add data.frames as tables, format is defined by the associated Word table style.
-   text: add text as paragraphs or inside an existing paragraph, format is defined by the associated Word paragraph and text styles.
-   field codes: add Word field codes inside paragraphs. Field codes is an old feature of MS Word to create calculated elements such as tables of content, automatic numbering and hyperlinks.

In a Word document, one can use cursor functions to reach the beginning of a document, its end or a particular paragraph containing a given text. This *cursor* concept has been implemented to make easier the post processing of files.

The file generation is performed with function `print`.

### import Word document in a data.frame

Function `docx_summary` read and import content of a Word document into a tibble object. The function handles paragraphs, tables and section breaks.

PowerPoint documents
--------------------

Function `read_pptx` will read an initial PowerPoint document (an empty one by default) and let you modify its content later.

The package provides functions to add R outputs into existing or new PowerPoint slides:

-   images: produce your plot in png or emf files and add them in a slide.
-   tables: add data.frames as tables, format is defined by the associated PowerPoint table style.
-   text: add text as paragraphs or inside an existing paragraph, format is defined in the corresponding layout of the slide.

In a PowerPoint document, one can set a slide as selected and reach a particular shape (and remove it or add text).

The file generation is performed with function `print`.

### import PowerPoint document in a data.frame

Function `pptx_summary` read and import content of a PowerPoint document into a tibble object. The function handles paragraphs, tables and images.

### Tables and package `flextable`

The package [flextable](https://github.com/davidgohel/flextable) brings a full API to produce nice tables and use them with `officer`.

Installation
------------

You can get the development version from GitHub:

``` r
devtools::install_github("davidgohel/officer")
```

Or the latest version on CRAN:

``` r
install.packages("officer")
```

Plan / to do list
-----------------

-   enable access to headers and footers in Word documents
-   improve tests and documentation (help is more than welcome)
