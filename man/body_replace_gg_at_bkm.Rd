% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/docx_replace.R
\name{body_replace_gg_at_bkm}
\alias{body_replace_gg_at_bkm}
\alias{body_replace_plot_at_bkm}
\title{Add plots at bookmark location in a 'Word' document}
\usage{
body_replace_gg_at_bkm(
  x,
  bookmark,
  value,
  width = 6,
  height = 5,
  res = 300,
  style = "Normal",
  scale = 1,
  keep = FALSE,
  ...
)

body_replace_plot_at_bkm(
  x,
  bookmark,
  value,
  width = 6,
  height = 5,
  res = 300,
  style = "Normal",
  keep = FALSE,
  ...
)
}
\arguments{
\item{x}{an rdocx object}

\item{bookmark}{bookmark id}

\item{value}{a ggplot object for body_replace_gg_at_bkm() or a set plot instructions
body_replace_plot_at_bkm(), see plot_instr().}

\item{width, height}{plot size in units expressed by the unit argument.
Defaults to a width of 6 and a height of 5 "in"ches.}

\item{res}{resolution of the png image in ppi}

\item{style}{paragraph style}

\item{scale}{Multiplicative scaling factor, same as in ggsave}

\item{keep}{Should the bookmark be preserved? Defaults to \code{FALSE},
i.e.the bookmark will be lost as a side effect.}

\item{...}{Arguments to be passed to png function.}
}
\description{
Use these functions if you want to replace a paragraph containing
a bookmark with a 'ggplot' or a base plot.
}
\examples{
if (require("ggplot2")) {
  gg_plot <- ggplot(data = iris) +
    geom_point(mapping = aes(Sepal.Length, Petal.Length))

  doc <- read_docx()
  doc <- body_add_par(doc, "insert_plot_here")
  doc <- body_bookmark(doc, "plot")
  doc <- body_replace_gg_at_bkm(doc, bookmark = "plot", value = gg_plot)
  print(doc, target = tempfile(fileext = ".docx"))
}
doc <- read_docx()
doc <- body_add_par(doc, "insert_plot_here")
doc <- body_bookmark(doc, "plot")
if (capabilities(what = "png")) {
  doc <- body_replace_plot_at_bkm(
    doc, bookmark = "plot",
    value = plot_instr(
      code = {
        barplot(1:5, col = 2:6)
      }
    )
  )
}
print(doc, target = tempfile(fileext = ".docx"))
}
