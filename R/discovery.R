#' List placeholders available on a layout/master
#'
#' @param template_path Path to an existing PowerPoint `.pptx` file.
#' @param layout Layout name.
#' @param master Optional master name for the layout. If `NULL`, first matching
#'   master is used.
#'
#' @return A data frame with placeholder labels, types, and dimensions.
#' @export
list_placeholders <- function(template_path, layout, master = NULL) {
  assert_single_string(template_path, "template_path")
  assert_single_string(layout, "layout")
  assert_optional_single_string(master, "master")

  if (!file.exists(template_path)) {
    stop("Template file does not exist: ", template_path, call. = FALSE)
  }

  x <- officer::read_pptx(path = template_path)
  resolved <- resolve_layout_master(x, layout = layout, master = master)
  props <- officer::layout_properties(
    x,
    layout = resolved$layout,
    master = resolved$master
  )

  keep <- c("type", "id", "ph_label", "ph", "offx", "offy", "cx", "cy")
  props[, intersect(keep, names(props)), drop = FALSE]
}
