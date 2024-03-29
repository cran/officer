% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/docx_replace.R
\name{sanitize_images}
\alias{sanitize_images}
\title{Remove unused media from a document}
\usage{
sanitize_images(x)
}
\arguments{
\item{x}{\code{rdocx} or \code{rpptx} object}
}
\description{
The function will scan the media
directory and delete images that are not used
anymore. This function is to be used when images
have been replaced many times.
}
\keyword{internal}
