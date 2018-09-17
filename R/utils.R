
#' Retrieve a data.frame by name from an environment
#'
#' @param df character, name of the object
#' @param env an environment
#'
#' @return the object
#' @noRd
#'
#' @importFrom utils data
#'
get_df <- function(df, env = globalenv()) {
  if (is.data.frame(df))
    return(df)
  if (df %in% ls(name = env)) {
    get(x = df, envir = env)
  } else {
    NULL
  }
}



#' Search for object with specific class in an environment
#'
#' @param what a class to look for
#' @param env An environment
#'
#' @return Character vector of the names of objects, NULL if none
#' @noRd
#'
#' @examples
#' \dontrun{
#'
#' search_obj("data.frame")
#'
#'
#' gg <- ggplot()
#' search_obj("ggplot")
#'
#' }
search_obj <- function(what = "data.frame", env = globalenv()) {
  all <- ls(name = env)
  objs <- lapply(
    X = all,
    FUN = function(x) {
      if (inherits(get(x, envir = env), what = what)) {
        x
      } else {
        NULL
      }
    }
  )
  objs <- unlist(objs)
  if (length(objs) == 1 && objs == "") {
    NULL
  } else {
    objs
  }
}
