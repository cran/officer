% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/formatting_properties.R
\name{fp_tabs}
\alias{fp_tabs}
\title{Tabs properties object}
\usage{
fp_tabs(...)
}
\arguments{
\item{...}{\link{fp_tab} objects}
}
\description{
create a set of tabulation mark properties object for Word or RTF.
Results can be used as arguments \code{tabs} of \code{\link[=fp_par]{fp_par()}} and will only have
effects in Word or RTF outputs.

Once a set of tabulation marks settings is defined, tabulation marks can
be added with \code{\link[=run_tab]{run_tab()}} inside a call to \code{\link[=fpar]{fpar()}} or
with \verb{\\t} within 'flextable' content.
}
\examples{
z <- fp_tabs(
  fp_tab(pos = 0.4, style = "decimal"),
  fp_tab(pos = 1, style = "decimal")
)
fpar(
  run_tab(), ftext("88."),
  run_tab(), ftext("987.45"),
  fp_p = fp_par(
    tabs = z
  )
)
}
\seealso{
Other functions for defining formatting properties: 
\code{\link{fp_border}()},
\code{\link{fp_cell}()},
\code{\link{fp_par}()},
\code{\link{fp_tab}()},
\code{\link{fp_text}()}
}
\concept{functions for defining formatting properties}
