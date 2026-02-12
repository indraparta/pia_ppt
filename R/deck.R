#' Open a PowerPoint template for editing
#'
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param include_existing_slides Whether to keep slides already present in the
#'   template. Defaults to `FALSE`, so only newly created slides are written.
#'
#' @return An `officer::rpptx` object.
#' @export
open_template <- function(cfg, include_existing_slides = FALSE) {
  assert_template_config(cfg)
  assert_single_logical(include_existing_slides, "include_existing_slides")

  x <- officer::read_pptx(path = cfg$template_path)
  if (isTRUE(include_existing_slides)) {
    return(x)
  }

  strip_all_slides(x)
}

#' Add a slide using a specific layout/master
#'
#' @param x An `officer::rpptx` object.
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param layout Layout name. Defaults to the configured primary content layout.
#' @param master Master name. If `NULL`, first matching master is used.
#'
#' @return Updated `officer::rpptx` object.
#' @export
add_master_slide <- function(
  x,
  cfg,
  layout = cfg$content_spec$layout,
  master = cfg$content_spec$master
) {
  assert_rpptx(x, "x")
  assert_template_config(cfg)
  assert_single_string(layout, "layout")
  assert_optional_single_string(master, "master")

  resolved <- resolve_layout_master(x, layout = layout, master = master)
  officer::add_slide(x, layout = resolved$layout, master = resolved$master)
}

#' Save a PowerPoint presentation
#'
#' @param x An `officer::rpptx` object.
#' @param target Output path for the generated `.pptx`.
#'
#' @return Invisibly returns `target`.
#' @export
write_presentation <- function(x, target) {
  assert_rpptx(x, "x")
  assert_single_string(target, "target")

  dir <- dirname(target)
  if (!dir.exists(dir)) {
    dir.create(dir, recursive = TRUE, showWarnings = FALSE)
  }

  print(x, target = target)
  invisible(target)
}
