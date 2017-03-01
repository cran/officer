## ---- echo = FALSE, message=FALSE, warning=FALSE-------------------------
dir.create("assets/docx", recursive = TRUE, showWarnings = FALSE)
office_doc_link <- function(url){
  stopifnot(requireNamespace("htmltools", quietly = TRUE))
  htmltools::tags$p(  htmltools::tags$span("Download file: "),
    htmltools::tags$a(basename(url), href = url))
}

## ------------------------------------------------------------------------
library(officer)
# Package `magrittr` makes officer usage easier.
library(magrittr)

## ------------------------------------------------------------------------
my_doc <- read_docx() 

# display styles
styles_info(my_doc)

## ------------------------------------------------------------------------
src <- tempfile(fileext = ".png")
png(filename = src, width = 5, height = 6, units = 'in', res = 300)
barplot(1:10, col = 1:10)
dev.off()

my_doc <- my_doc %>% 
  body_add_img(src = src, width = 5, height = 6, style = "centered")

## ------------------------------------------------------------------------
my_doc <- my_doc %>% 
  body_add_par("Hello world!", style = "Normal") %>% 
  body_add_par("", style = "Normal") # blank paragraph

## ------------------------------------------------------------------------
my_doc <- my_doc %>% 
  body_add_table(iris, style = "table_template")

## ------------------------------------------------------------------------
print(my_doc, target = "assets/docx/first_example.docx") %>% 
  invisible()

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/officer/articles/", "assets/docx/first_example.docx" ) )

## ------------------------------------------------------------------------
read_docx() %>%
  body_add_par("paragraph 1", style = "Normal") %>%
  body_add_par("paragraph 2", style = "Normal") %>%
  body_add_par("paragraph 3", style = "Normal") %>%
  body_add_par("paragraph 4", style = "Normal") %>%
  body_add_par("paragraph 5", style = "Normal") %>%
  body_add_par("paragraph 6", style = "Normal") %>%
  body_add_par("paragraph 7", style = "Normal") %>%
  print(target = "assets/docx/init_doc.docx" ) %>% 
  invisible()

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/officer/articles/", "assets/docx/init_doc.docx" ) )

## ------------------------------------------------------------------------
doc <- read_docx(path = "assets/docx/init_doc.docx") %>%

  # default template contains only an empty paragraph
  # Using cursor_begin and body_remove, we can delete it
  cursor_begin() %>% body_remove() %>%

  # Let add text at the beginning of the
  # paragraph containing text "paragraph 4"
  cursor_reach(keyword = "paragraph 4") %>%
  slip_in_text("This is ", pos = "before", style = "Default Paragraph Font") %>%

  # move the cursor forward and end a section
  cursor_forward() %>%
  body_add_par("The section stop here", style = "Normal") %>%
  body_end_section(landscape = TRUE) %>%

  # move the cursor at the end of the document
  cursor_end() %>%
  body_add_par("The document ends now", style = "Normal")

print(doc, target = "assets/docx/cursor.docx") %>% 
  invisible()

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/officer/articles/", "assets/docx/cursor.docx" ) )

## ------------------------------------------------------------------------
library(magrittr)
img.file <- file.path( Sys.getenv("R_HOME"), "doc", "html", "logo.jpg" )

read_docx() %>%
  body_add_par("R logo: ", style = "Normal") %>%
  slip_in_img(src = img.file, style = "strong", width = .3, height = .3) %>%
  slip_in_text("This is official ", style = "strong", pos = "before") %>%
  slip_in_text(" that can be found here: ", style = "strong", pos = "after") %>%
  slip_in_text(img.file, style = "strong", pos = "after") %>%
  print(target = "assets/docx/slip_in_demo.docx") %>% 
  invisible()

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/officer/articles/", "assets/docx/slip_in_demo.docx" ) )

## ------------------------------------------------------------------------
library(officer)
library(magrittr)

str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>% 
  rep(20) %>% paste(collapse = "")
str2 <- "Drop that text" 
str3 <- "Aenean venenatis varius elit et fermentum vivamus vehicula. " %>% 
  rep(20) %>% paste(collapse = "")

my_doc <- read_docx()  %>% 
  body_add_par(value = str1, style = "Normal") %>% 
  body_add_par(value = str2, style = "centered") %>% 
  body_add_par(value = str3, style = "Normal") 

print(my_doc, target = "assets/docx/ipsum_doc.docx") %>% invisible()

## ------------------------------------------------------------------------
my_doc <- read_docx(path = "assets/docx/ipsum_doc.docx")  %>% 
  cursor_reach(keyword = "that text") %>% 
  body_remove()

print(my_doc, target = "assets/docx/ipsum_doc.docx") %>% invisible()

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/officer/articles/", "assets/docx/ipsum_doc.docx" ) )

## ------------------------------------------------------------------------
str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>% 
  rep(30) %>% paste(collapse = "")
str2 <- "Aenean venenatis varius elit et fermentum vivamus vehicula. " %>% 
  rep(30) %>% paste(collapse = "")

my_doc <- read_docx()  %>% 
  slip_in_text(str = str1, style = "strong") %>% 
  body_add_par(value = str2, style = "centered") %>% 
  body_end_section(landscape = TRUE, colwidths = c(.6, .4), space = .05, sep = FALSE) %>%
  body_add_par(value = str3, style = "Normal") 
print(my_doc, target = "assets/docx/section.docx") %>% invisible()

## ----echo=FALSE----------------------------------------------------------
office_doc_link( url = paste0( "https://davidgohel.github.io/officer/articles/", "assets/docx/section.docx" ) )

