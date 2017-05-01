## ------------------------------------------------------------------------
library(officer)
example_docx <- system.file(package = "officer", "doc_examples/example.docx")
doc <- read_docx(example_docx)
content <- docx_summary(doc)
content

## ---- message=FALSE, warning=FALSE---------------------------------------
library(dplyr)
content %>% group_by(content_type) %>% summarise(n = n_distinct(doc_index))

## ---- message=FALSE, warning=FALSE---------------------------------------
par_data <- content %>% filter(content_type %in% "paragraph") %>% 
  select(doc_index, style_name, text, level, num_id) %>% 
  # let's make text shorter so it can be display in that vignette
  mutate(text = substr(text, start = 1, 
                       stop = ifelse(nchar(text)<30, nchar(text), 30) ))

par_data

## ---- message=FALSE, warning=FALSE---------------------------------------
table_cells <- content %>% filter(content_type %in% "table cell")
print(table_cells)

## ------------------------------------------------------------------------
table_body <- table_cells %>% 
  filter(!is_header) %>% 
  select(row_id, cell_id, text)
table_body

## ------------------------------------------------------------------------
if( require("tidyr"))
  table_body %>% spread(cell_id, text)  

## ------------------------------------------------------------------------
if( require("tidyr"))
  table_cells %>% 
    filter(is_header) %>% 
    select(row_id, cell_id, text) %>% 
    spread(cell_id, text)  

## ------------------------------------------------------------------------
example_pptx <- system.file(package = "officer", "doc_examples/example.pptx")
doc <- read_pptx(example_pptx)
content <- pptx_summary(doc)
content

## ---- message=FALSE, warning=FALSE---------------------------------------
content %>% group_by(content_type) %>% summarise(n = n_distinct(id))

## ---- message=FALSE, warning=FALSE---------------------------------------
par_data <- content %>% filter(content_type %in% "paragraph") %>% 
  select(id, text)

par_data

## ------------------------------------------------------------------------
image_row <- content %>% filter(content_type %in% "image")
media_extract(doc, path = image_row$media_file, target = "extract.png")

## ---- message=FALSE, warning=FALSE---------------------------------------
table_cells <- content %>% filter(content_type %in% "table cell")
table_cells

## ------------------------------------------------------------------------
if( require("tidyr"))
  table_cells %>% 
    select(row_id, cell_id, text) %>% 
    spread(cell_id, text)  

