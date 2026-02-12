#' Add a front-page title slide and move it to slide 1
#'
#' @param x An `officer::rpptx` object.
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param title A title string for the front slide.
#' @param subtitle Optional subtitle string.
#'
#' @return Updated `officer::rpptx` object.
#' @export
add_front_slide <- function(x, cfg, title, subtitle = NULL) {
  assert_rpptx(x, "x")
  assert_template_config(cfg)
  assert_single_string(title, "title")
  assert_optional_single_string(subtitle, "subtitle")

  spec <- cfg$front_spec
  resolved <- resolve_layout_master(x, layout = spec$layout, master = spec$master)
  assert_placeholders_exist(
    x = x,
    layout = resolved$layout,
    master = resolved$master,
    labels = unname(unlist(spec$placeholders[c("title", "subtitle")])),
    context = "front_spec"
  )

  x <- officer::add_slide(x, layout = resolved$layout, master = resolved$master)
  x <- inject_slide_title(
    x = x,
    title = title,
    cfg = cfg,
    ph_label = spec$placeholders$title
  )

  if (!is.null(subtitle)) {
    x <- officer::ph_with(
      x = x,
      value = subtitle,
      location = officer::ph_location_label(ph_label = spec$placeholders$subtitle)
    )
  }

  idx <- slide_count(x)
  if (idx > 1L) {
    x <- move_slide_to_start(x, from_index = idx)
  }

  x
}
