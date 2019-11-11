officer R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->

The officer package lets R users manipulate Word (`.docx`) and
PowerPoint (`*.pptx`) documents. In short, one can add images, tables
and text into documents from R. An initial document can be provided;
contents, styles and properties of the original document will then be
available.

<img src="https://www.ardata.fr/img/hexbin/officer.svg" width = "15%"/>

> Make corporate reporting with minimum hassle

[![Travis-CI Build
Status](https://travis-ci.org/davidgohel/officer.svg?branch=master)](https://travis-ci.org/davidgohel/officer)
[![Build
status](https://ci.appveyor.com/api/projects/status/github/davidgohel/officer?branch=master&svg=true)](https://ci.appveyor.com/project/davidgohel/officer/branch/master)
[![version](http://www.r-pkg.org/badges/version/officer)](https://CRAN.R-project.org/package=officer)
![cranlogs](http://cranlogs.r-pkg.org./badges/officer)
[![codecov](https://codecov.io/gh/davidgohel/officer/branch/master/graph/badge.svg)](https://codecov.io/gh/davidgohel/officer)
![Active](http://www.repostatus.org/badges/latest/active.svg)

## Word documents

The `read_docx()` function will read an initial Word document (an empty
one by default) and lets you modify its content later.

The package provides functions to add R outputs into a Word document:

  - images: produce your plot in png or emf files and add them into the
    document, as a whole paragraph or inside a paragraph.
  - tables: add data.frames as tables, format is defined by the
    associated Word table style.
  - text: add text as paragraphs or inside an existing paragraph, format
    is defined by the associated Word paragraph and text styles.
  - field codes: add Word field codes inside paragraphs. Field codes is
    an old feature of MS Word to create calculated elements such as
    tables of contents, automatic numbering and hyperlinks.

In a Word document, one can use cursor functions to reach the beginning
or end of a document, or a particular paragraph containing a given text.
This *cursor* concept has been implemented to make the post processing
of files easier.

File generation is performed with the `print` function.

### import Word document in a data.frame

The function `docx_summary()` reads and imports content of a Word
document into a data.frame. The function handles paragraphs, tables and
section breaks.

## PowerPoint documents

The function `read_pptx()` will read an initial PowerPoint document (an
empty one by default) and let you modify its content later.

The package provides functions to add R outputs into existing or new
PowerPoint slides:

  - images: produce your plot in png or emf files and add them in a
    slide.
  - tables: add data.frames as tables, format is defined by the
    associated PowerPoint table style.
  - text: add text as paragraphs or inside an existing paragraph, format
    is defined in the corresponding layout of the slide.

In a PowerPoint document, one can set a slide as selected and reach a
particular shape (and remove it or add text).

File generation is performed with the `print()` function.

### import PowerPoint document in a data.frame

The `pptx_summary()` function reads and imports content of a PowerPoint
document into a tibble object. The function handles paragraphs, tables
and images.

## Extensions

### Tables and package `flextable`

The package [flextable](https://github.com/davidgohel/flextable) brings
a full API to produce nice tables and use them with `officer`. Tables
can be written in PowerPoint documents and Word documents. An option is
available to render flextables in rmarkdown (HTML and Word outputs).

### Vector graphics with package `rvg`

The package [rvg](https://github.com/davidgohel/rvg) brings an API to
produce nice vector graphics that can be embedded in PowerPoint
documents or Excel workbooks with `officer`.

### Native office charts with package `mschart`

The package [mschart](https://github.com/ardata-fr/mschart) combined
with `officer` can produce native office charts in PowerPoint and Word
documents.

## Installation

You can get the development version from GitHub:

``` r
devtools::install_github("davidgohel/officer")
```

Or the latest version on CRAN:

``` r
install.packages("officer")
```

## Resources

### Online documentation

The help pages are located at <https://davidgohel.github.io/officer/>.

### Getting help

If you have questions about how to use the package, [visit Stack
Overflow’s `officer`
tag](http://stackoverflow.com/questions/tagged/officer) and post your
question there. I usually read them and answer when possible.

## Contributing to the package

### Code of Conduct

Anyone getting involved in this package agrees to our [Code of
Conduct](https://github.com/davidgohel/officer/blob/master/CONDUCT.md).

### Bug reports

When you file a [bug
report](https://github.com/davidgohel/officer/issues), please spend some
time making it easy for me to follow and reproduce. The more time you
spend on making the bug report coherent, the more time I can dedicate to
investigate the bug as opposed to the bug report.

### Contributing to the package development

A great way to start is to contribute an example or improve the
documentation.

If you want to submit a Pull Request to integrate functions of yours,
please provide:

  - the new function(s) with code and roxygen tags (with examples)
  - a new section in the appropriate vignette that describes how to use
    the new function
  - add corresponding tests in directory `tests/testthat`.

By using rhub (run `rhub::check_for_cran()`), you will see if everything
is ok. When submitted, the PR will be evaluated automatically on travis
and appveyor and you will be able to see if something broke.
