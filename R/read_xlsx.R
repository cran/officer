# worksheets ------------------------------------------------------------

worksheets <- R6Class(
  "worksheets",
  inherit = openxml_document,

  public = list(

    initialize = function( path ) {
      super$initialize(character(0))
      private$package_dir <- path
      presentation_filename <- file.path(path, "xl", "workbook.xml")
      self$feed(presentation_filename)

      slide_df <- self$get_sheets_df()
      private$sheet_id <- slide_df$sheet_id
      private$sheet_rid <- slide_df$rid
      private$sheet_name <- slide_df$name
    },

    view_on_sheet = function(sheet){
      sheet_id <- self$get_sheet_id(sheet)
      wb_view <- xml_find_first(self$get(), "d1:bookViews/d1:workbookView")
      xml_attr(wb_view, "activeTab") <- sheet_id - 1
      self$save()
    },

    add_sheet = function(target, label){

      private$rels_doc$add(id = paste0("rId", private$rels_doc$get_next_id() ),
                           type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                           target = target )
      rels <- private$rels_doc$get_data()
      rid <- rels[rels$target %in% target, "id"]

      ids <- private$sheet_id
      new_id <- max(ids) + 1

      private$sheet_id <- c( private$sheet_id, new_id)
      private$sheet_rid <- c( private$sheet_rid, rid)
      private$sheet_name <- c( private$sheet_name, label)

      children_ <- xml_children(self$get())
      sheets_id <- which(sapply(children_, function(x) xml_name(x)=="sheets" ))
      xml_list <- xml_children(children_[[sheets_id]])

      xml_elt <- paste(
        sprintf("<sheet name=\"%s\" sheetId=\"%.0f\" r:id=\"%s\"/>",
                htmlEscapeCopy(private$sheet_name), private$sheet_id, private$sheet_rid),
        collapse = "" )
      xml_elt <- paste0( "<sheets xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
                                     xml_elt, "</sheets>")
      xml_elt <- as_xml_document(xml_elt)

      if( !inherits(xml_list, "xml_missing")){
        xml_replace(children_[[sheets_id]], xml_elt)
      } else{
        stop("could not find sheets entity")
      }

      self
    },

    get_new_sheetname = function(){
      sheet_dir <- file.path(private$package_dir, "xl/worksheets")
      if( !file.exists(sheet_dir)){
        dir.create(file.path(sheet_dir, "_rels"), showWarnings = FALSE, recursive = TRUE)
      }

      sheet_files <- basename( list.files(sheet_dir, pattern = "\\.xml$") )
      sheet_name <- "sheet1.xml"
      if( length(sheet_files)){
        slide_index <- as.integer(gsub("^(sheet)([0-9]+)(\\.xml)$", "\\2", sheet_files ))
        sheet_name <- gsub(pattern = "[0-9]+", replacement = max(slide_index) + 1, sheet_name)
      }
      sheet_name
    },

    sheet_names = function( ){
      private$sheet_name
    },

    get_sheet_id = function(name){
      sheets_df <- self$get_sheets_df()
      bool_name_in_list <- sheets_df$name %in% name
      n_matches <- sum(bool_name_in_list, na.rm = TRUE)
      if(n_matches < 1 )
        stop("could not find ", shQuote(name), " sheet", call. = FALSE)
      sheets_df$sheet_id[bool_name_in_list]
    },

    get_sheets_df = function() {
      children_ <- xml_children(self$get())
      sheets_id <- which(sapply(children_, function(x) xml_name(x)=="sheets" ))
      sheet_nodes <- xml_children(children_[[sheets_id]])
      data.frame(stringsAsFactors = FALSE,
                 name = xml_attr(sheet_nodes, "name"),
              sheet_id = as.integer(xml_attr(sheet_nodes, "sheetId")),
              rid = xml_attr(sheet_nodes, "id") )
    }




  ),
  private = list(

    sheet_id = NULL,
    sheet_rid = NULL,
    sheet_name = NULL,
    package_dir = NULL

  )
)

# sheet ------------------------------------------------------------
sheet <- R6Class(
  "sheet",
  inherit = openxml_document,
  public = list(

    feed = function( file ) {
      private$filename <- file
      private$doc <- read_xml(file)

      private$rels_filename <- file.path( dirname(file), "_rels", paste0(basename(file), ".rels") )

      if( file.exists(private$rels_filename) ){
        private$rels_doc <- relationship$new()$feed_from_xml(private$rels_filename)
      } else {
        new_rel <- relationship$new()
        new_rel$write(private$rels_filename)
        private$rels_doc <- new_rel
      }
      self
    }

  )

)


# dir_sheets ----
dir_sheet <- R6Class(
  "dir_sheet",
  public = list(

    initialize = function( x ) {
      private$package_dir <- x$package_dir
      private$sheets_path <- "xl/worksheets"
      self$update()
    },

    update = function(){
      dir_ <- file.path(private$package_dir, private$sheets_path)
      filenames <- list.files(path = dir_, pattern = "\\.xml$", full.names = TRUE)

      private$collection <- lapply( filenames, function(x, path){
        sheet$new(path)$feed(x)
      }, private$sheets_path)

      names(private$collection) <- sapply(private$collection, function(x) x$name() )
      self
    },

    get_sheet_list = function(){
      dir_ <- file.path(private$package_dir, private$sheets_path)
      filenames <- list.files(path = dir_, pattern = "\\.xml$", full.names = TRUE)
      sheet_index <- seq_along(filenames)
      if( length(filenames)){
        filenames <- basename( filenames )
        sheet_index <- as.integer(gsub("^(sheet)([0-9]+)(\\.xml)$", "\\2", filenames ))
        filenames <- filenames[order(sheet_index)]
      }
      filenames
    },

    get_sheet = function(id){

      l_ <- self$length()
      if( is.null(id) || !between(id, 1, l_ ) ){
        stop("unvalid id ", id, " (", l_," sheet(s))", call. = FALSE)
      }
      sheet_files <- self$get_sheet_list()
      index <- which( names(private$collection) == sheet_files[id])
      private$collection[[index]]
    },

    length = function(){
      length(private$collection)
    }
  ),
  private = list(
    collection = NULL,
    package_dir = NULL,
    sheets_path = NULL
  )
)



# read_xlsx ----
#' @export
#' @title Create an 'Excel' document object
#' @description Read and import an xlsx file as an R object
#' representing the document. This function is experimental.
#' @param path path to the xlsx file to use as base document.
#' @param x an rxlsx object
#' @examples
#' read_xlsx()
read_xlsx <- function( path = NULL ){

  if( !is.null(path) && !file.exists(path))
    stop("could not find file ", shQuote(path), call. = FALSE)

  if( is.null(path) )
    path <- system.file(package = "officer", "template/template.xlsx")

  if(!grepl("\\.xlsx$", path, ignore.case = TRUE)){
    stop("read_xlsx only support xlsx files", call. = FALSE)
  }

  package_dir <- tempfile()
  unpack_folder( file = path, folder = package_dir )

  obj <- structure(list( package_dir = package_dir ),
                   .Names = c("package_dir"),
                   class = "rxlsx")

  obj$content_type <- content_type$new( package_dir )
  obj$worksheets <- worksheets$new(package_dir)
  obj$sheets <- dir_sheet$new( obj )
  obj$core_properties <- read_core_properties(obj$package_dir)

  obj
}

#' @export
#' @title Add a sheet
#' @description Add a sheet into an xlsx worksheet.
#' @param x rxlsx object
#' @param label sheet label
#' @examples
#' my_ws <- read_xlsx()
#' my_pres <- add_sheet(my_ws, label = "new sheet")
add_sheet <- function( x, label ){

  if(label %in% x$worksheets$sheet_names()){
    stop("sheet ", shQuote(label), " already exist")
  }

  new_slidename <- x$worksheets$get_new_sheetname()

  xml_file <- file.path(x$package_dir, "xl/worksheets", new_slidename)

  template_file <- system.file(package = "officer", "template/sheet.xml")
  file.copy(template_file, xml_file, copy.mode = FALSE)

  rel_filename <- file.path(
    dirname(xml_file), "_rels",
    paste0(basename(xml_file), ".rels") )
  dir.create(dirname(rel_filename), showWarnings = FALSE)
  template_rel_file <- system.file(package = "officer", "template/sheet.xml.rels")
  file.copy(template_rel_file, rel_filename, copy.mode = FALSE)

  # update presentation elements
  x$worksheets$add_sheet(target = file.path( "worksheets", new_slidename), label = label )

  partname <- file.path( "/xl/worksheets", new_slidename )
  override <- setNames("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", partname )
  x$content_type$add_override(value = override)

  x$sheets$update()

  sheet_select(x, sheet = label)
}

#' @export
#' @rdname read_xlsx
length.rxlsx <- function( x ){
  x$sheets$length()
}

#' @export
#' @title Select sheet
#' @description Set a particular sheet selected when workbook will be
#' edited.
#' @param x rxlsx object
#' @param sheet sheet name
#' @examples
#' my_ws <- read_xlsx()
#' my_pres <- add_sheet(my_ws, label = "new sheet")
#' my_pres <- sheet_select(my_ws, sheet = "new sheet")
#' print(my_ws, target = tempfile(fileext = ".xlsx") )
sheet_select <- function( x, sheet ){
  x$worksheets$view_on_sheet(sheet)
  x
}

#' @export
#' @param target path to the xlsx file to write
#' @param ... unused
#' @rdname read_xlsx
#' @examples
#' x <- read_xlsx()
#' print(x, target = tempfile(fileext = ".xlsx"))
print.rxlsx <- function(x, target = NULL, ...){

  if( is.null( target) ){
    cat("xlsx document with", length(x), "sheet(s):\n")
    print( x$worksheets$sheet_names() )
    return(invisible())
  }

  if( !grepl(x = target, pattern = "\\.(xlsx)$", ignore.case = TRUE) )
    stop(target , " should have '.xlsx' extension.")

  x$worksheets$save()
  x$content_type$save()

  x$core_properties['modified','value'] <- format( Sys.time(), "%Y-%m-%dT%H:%M:%SZ")
  x$core_properties['lastModifiedBy','value'] <- Sys.getenv("USER")
  write_core_properties(x$core_properties, x$package_dir)

  invisible(pack_folder(folder = x$package_dir, target = target ))
}
