

#' Open a data.frame in 'Excel'
#'
#' @param df The name of a \code{data.frame} or a \code{data.frame} itself, if not provided, an addin is launched.
#'
#' @export
#'
#' @examples
#' \dontrun{
#'
#' # Launch addin
#' view_in_xl()
#'
#' # or pass the name of a data.frame
#' view_in_xl("iris")
#' 
#' # or directly the object
#' view_in_xl(iris)
#'
#' }
#' @importFrom miniUI miniPage miniButtonBlock miniContentPanel
#' @importFrom rstudioapi getActiveDocumentContext
#' @importFrom shiny actionLink actionButton icon observeEvent
#'  runGadget dialogViewer tags stopApp
#' @importFrom utils browseURL
#' @importFrom writexl write_xlsx
view_in_xl <- function(df = NULL) {
  if (is.null(df)) {
    context <- getActiveDocumentContext()
    df <- context$selection[[1]]$text
    is_df <- tryCatch({
      test <- get(df, envir = .GlobalEnv)
      test <- as.data.frame(test)
      list(res = is.data.frame(test))
    }, error = function(e) {
      list(res = FALSE)
    })
    if (!is_df$res) {
      df <- search_obj()
    }
  } else if (is.data.frame(df)) {
    df <- deparse(substitute(df))
  }
  if (length(df) == 0) {
    message("It seems that there are no data.frames in global environment...")
    return(invisible())
  }
  if (length(df) == 1) {
    tmp <- tempfile(fileext = ".xlsx")
    df <- get_df(df)
    if (is.null(df)) {
      message("Selected object must be a data.frame in GlobalEnv.")
      return(invisible())
    }
    df <- as.data.frame(df)
    write_xlsx(x = df, path = tmp)
    browseURL(url = tmp)
    return(invisible(tmp))
  } else {
    tags_df <- lapply(
      X = df,
      FUN = function(x) {
        obj <- get_df(x)
        tags$li(
          actionLink(
            inputId = paste0("view_in_xl_", x),
            label = sprintf("%s (%s obs. of  %s variables)", x, nrow(obj), ncol(obj))
          )
        )
      }
    )
    ui <- miniPage(
      miniContentPanel(
        tags$ul(tags_df)
      ),
      miniButtonBlock(
        actionButton(
          inputId = "close", label = "Close",
          icon = icon("remove"),
          class = "btn-block btn-primary"
        )
      )
    )
    server <- function(input, output, session) {
      lapply(
        X = paste0("view_in_xl_", df),
        FUN = function(x) {
          observeEvent(input[[x]], {
            tmp <- tempfile(fileext = ".xlsx")
            obj <- gsub(pattern = "view_in_xl_", replacement = "", x = x)
            write_xlsx(x = get_df(obj), path = tmp)
            browseURL(url = tmp)
          }, ignoreInit = TRUE)
        }
      )
      observeEvent(input$close, stopApp())
    }
    inviewer <- dialogViewer(
      "View a data.frame in Excel",
      width = 450, height = 180
    )
    runGadget(app = ui, server = server, viewer = inviewer)
  }
}

