% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/rtf.R
\name{rtf_add}
\alias{rtf_add}
\alias{rtf_add.block_section}
\alias{rtf_add.character}
\alias{rtf_add.factor}
\alias{rtf_add.double}
\alias{rtf_add.fpar}
\alias{rtf_add.block_list}
\alias{rtf_add.gg}
\alias{rtf_add.plot_instr}
\title{Add content into an RTF document}
\usage{
rtf_add(x, value, ...)

\method{rtf_add}{block_section}(x, value, ...)

\method{rtf_add}{character}(x, value, ...)

\method{rtf_add}{factor}(x, value, ...)

\method{rtf_add}{double}(x, value, formatter = formatC, ...)

\method{rtf_add}{fpar}(x, value, ...)

\method{rtf_add}{block_list}(x, value, ...)

\method{rtf_add}{gg}(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
)

\method{rtf_add}{plot_instr}(
  x,
  value,
  width = 6,
  height = 5,
  res = 300,
  scale = 1,
  ppr = fp_par(text.align = "center"),
  ...
)
}
\arguments{
\item{x}{rtf object, created by \code{\link[=rtf_doc]{rtf_doc()}}.}

\item{value}{object to add in the document. Supported objects
are vectors, graphics, block of formatted paragraphs. Use package
'flextable' to add tables.}

\item{...}{further arguments passed to or from other methods. When
adding a \code{ggplot} object or \link{plot_instr}, these arguments will be used
by png function. See section 'Methods' to see what arguments can be used.}

\item{formatter}{function used to format the numerical values}

\item{width}{height in inches}

\item{height}{height in inches}

\item{res}{resolution of the png image in ppi}

\item{scale}{Multiplicative scaling factor, same as in ggsave}

\item{ppr}{\code{\link[=fp_par]{fp_par()}} to apply to paragraph.}
}
\description{
This function add 'officer' objects into an RTF document.
Values are added as new paragraphs. See section 'Methods (by class)'
that list supported objects.
}
\section{Methods (by class)}{
\itemize{
\item \code{rtf_add(block_section)}: add a new section definition

\item \code{rtf_add(character)}: add characters as new paragraphs

\item \code{rtf_add(factor)}: add a factor vector as new paragraphs

\item \code{rtf_add(double)}: add a double vector as new paragraphs

\item \code{rtf_add(fpar)}: add an \code{\link[=fpar]{fpar()}}

\item \code{rtf_add(block_list)}: add an \code{\link[=block_list]{block_list()}}

\item \code{rtf_add(gg)}: add a ggplot2

\item \code{rtf_add(plot_instr)}: add a \code{\link[=plot_instr]{plot_instr()}} object

}}
\examples{
library(officer)

def_text <- fp_text_lite(color = "#006699", bold = TRUE)
center_par <- fp_par(text.align = "center", padding = 3)

doc <- rtf_doc(
  normal_par = fp_par(line_spacing = 1.4, padding = 3)
)

doc <- rtf_add(
  x = doc,
  value = fpar(
    ftext("how are you?", prop = def_text),
    fp_p = fp_par(text.align = "center")
  )
)

a_paragraph <- fpar(
  ftext("Here is a date: ", prop = def_text),
  run_word_field(field = "Date \\\\@ \"MMMM d yyyy\""),
  fp_p = center_par
)
doc <- rtf_add(
  x = doc,
  value = block_list(
    a_paragraph,
    a_paragraph,
    a_paragraph
  )
)

if (require("ggplot2")) {
  gg <- gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))
  doc <- rtf_add(doc, gg,
    width = 3, height = 4,
    ppr = center_par
  )
}
anyplot <- plot_instr(code = {
  barplot(1:5, col = 2:6)
})
doc <- rtf_add(doc, anyplot,
  width = 5, height = 4,
  ppr = center_par
)

print(doc, target = tempfile(fileext = ".rtf"))
}
