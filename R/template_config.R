#' Create a template configuration for PIA slide generation
#'
#' @param template_path Optional path to an existing PowerPoint `.pptx` file.
#'   If `NULL`, uses the bundled package template
#'   `inst/templates/pia_slidemaster.pptx`.
#' @param primary_layout Layout used for the primary content slide.
#' @param primary_master Optional master for `primary_layout`. If `NULL`, the
#'   first matching master is used.
#' @param placeholders Named list mapping content placeholders. Required keys:
#'   `title`, `chart`, `bullets`.
#' @param front_layout Layout used for the front title slide.
#' @param front_master Optional master for `front_layout`. If `NULL`, the first
#'   matching master is used.
#' @param front_placeholders Named list mapping front-slide placeholders.
#'   Required keys: `title`, `subtitle`.
#'
#' @return A `piappt_template` object.
#' @export
new_pia_template <- function(
  template_path = NULL,
  primary_layout = "03_Text-Chart-2",
  primary_master = NULL,
  placeholders = list(
    title = "Title 1",
    chart = "Content Placeholder 12",
    bullets = "Text Placeholder 9"
  ),
  front_layout = "01_Title",
  front_master = NULL,
  front_placeholders = list(
    title = "Title 1",
    subtitle = "Subtitle 2"
  )
) {
  assert_optional_single_string(template_path, "template_path")
  assert_single_string(primary_layout, "primary_layout")
  assert_optional_single_string(primary_master, "primary_master")
  assert_single_string(front_layout, "front_layout")
  assert_optional_single_string(front_master, "front_master")

  template_path <- resolve_template_path(template_path)

  content_placeholders <- validate_placeholder_map(
    map = placeholders,
    required_names = c("title", "chart", "bullets"),
    arg_name = "placeholders"
  )

  front_placeholders <- validate_placeholder_map(
    map = front_placeholders,
    required_names = c("title", "subtitle"),
    arg_name = "front_placeholders"
  )

  cfg <- list(
    template_path = normalizePath(template_path, winslash = "/", mustWork = TRUE),
    content_spec = list(
      layout = primary_layout,
      master = primary_master,
      placeholders = content_placeholders
    ),
    front_spec = list(
      layout = front_layout,
      master = front_master,
      placeholders = front_placeholders
    )
  )

  class(cfg) <- c("piappt_template", "list")
  cfg
}

resolve_template_path <- function(template_path = NULL) {
  if (!is.null(template_path)) {
    if (!file.exists(template_path)) {
      stop("Template file does not exist: ", template_path, call. = FALSE)
    }
    return(template_path)
  }

  bundled <- system.file(
    "templates",
    "pia_slidemaster.pptx",
    package = "piappt"
  )

  if (nzchar(bundled) && file.exists(bundled)) {
    return(bundled)
  }

  local_candidates <- c(
    file.path(getwd(), "inst", "templates", "pia_slidemaster.pptx"),
    file.path(getwd(), "pia_slidemaster.pptx")
  )
  local_candidates <- local_candidates[file.exists(local_candidates)]
  if (length(local_candidates) > 0) {
    return(local_candidates[[1]])
  }

  stop(
    "No template path provided and bundled template was not found. ",
    "Pass `template_path` explicitly or install the package with ",
    "`inst/templates/pia_slidemaster.pptx` included.",
    call. = FALSE
  )
}
