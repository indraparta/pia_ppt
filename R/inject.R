#' Inject a chart into the current slide
#'
#' @param x An `officer::rpptx` object.
#' @param chart A chart-like object supported by `officer::ph_with()`.
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param ph_label Placeholder label to target. Defaults to configured content
#'   chart placeholder.
#' @param location Optional explicit `officer::ph_location_*()` object.
#'
#' @return Updated `officer::rpptx` object.
#' @export
inject_chart <- function(
  x,
  chart,
  cfg,
  ph_label = cfg$content_spec$placeholders$chart,
  location = NULL
) {
  assert_rpptx(x, "x")
  assert_template_config(cfg)

  if (missing(chart) || is.null(chart)) {
    stop("`chart` must be provided.", call. = FALSE)
  }

  loc <- resolve_location(ph_label = ph_label, location = location)
  officer::ph_with(x = x, value = chart, location = loc)
}

#' Inject title text into the current slide
#'
#' @param x An `officer::rpptx` object.
#' @param title A single title string.
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param ph_label Placeholder label to target. Defaults to configured content
#'   title placeholder.
#' @param location Optional explicit `officer::ph_location_*()` object.
#'
#' @return Updated `officer::rpptx` object.
#' @export
inject_slide_title <- function(
  x,
  title,
  cfg,
  ph_label = cfg$content_spec$placeholders$title,
  location = NULL
) {
  assert_rpptx(x, "x")
  assert_template_config(cfg)
  assert_single_string(title, "title")

  loc <- resolve_location(ph_label = ph_label, location = location)
  officer::ph_with(x = x, value = title, location = loc)
}

#' Inject dot points into the current slide
#'
#' @param x An `officer::rpptx` object.
#' @param bullets Character vector of bullet text.
#' @param levels Optional integer vector of bullet indentation levels.
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param ph_label Placeholder label to target. Defaults to configured content
#'   bullet placeholder.
#' @param location Optional explicit `officer::ph_location_*()` object.
#'
#' @return Updated `officer::rpptx` object.
#' @export
inject_dotpoints <- function(
  x,
  bullets,
  levels = NULL,
  cfg,
  ph_label = cfg$content_spec$placeholders$bullets,
  location = NULL
) {
  assert_rpptx(x, "x")
  assert_template_config(cfg)

  if (missing(bullets) || length(bullets) == 0) {
    stop("`bullets` must contain at least one item.", call. = FALSE)
  }

  bullets <- as.character(bullets)

  if (is.null(levels)) {
    levels <- rep.int(1L, length(bullets))
  }

  if (!is.numeric(levels) || length(levels) != length(bullets) ||
      any(is.na(levels)) || any(levels < 1)) {
    stop("`levels` must be positive integers with one value per bullet.", call. = FALSE)
  }

  loc <- resolve_location(ph_label = ph_label, location = location)
  value <- officer::unordered_list(
    str_list = bullets,
    level_list = as.integer(levels)
  )

  officer::ph_with(x = x, value = value, location = loc)
}

#' Add a primary content slide
#'
#' @param x An `officer::rpptx` object.
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param chart A chart-like object supported by `officer::ph_with()`.
#' @param title A title string for the content slide.
#' @param bullets Character vector of bullet text.
#' @param bullet_levels Optional integer vector of bullet indentation levels.
#'
#' @return Updated `officer::rpptx` object.
#' @export
add_primary_slide <- function(
  x,
  cfg,
  chart,
  title,
  bullets,
  bullet_levels = NULL
) {
  assert_rpptx(x, "x")
  assert_template_config(cfg)
  assert_single_string(title, "title")

  spec <- cfg$content_spec
  resolved <- resolve_layout_master(x, layout = spec$layout, master = spec$master)
  assert_placeholders_exist(
    x = x,
    layout = resolved$layout,
    master = resolved$master,
    labels = unname(unlist(spec$placeholders[c("title", "chart", "bullets")])),
    context = "content_spec"
  )

  x <- officer::add_slide(x, layout = resolved$layout, master = resolved$master)
  x <- inject_chart(x = x, chart = chart, cfg = cfg, ph_label = spec$placeholders$chart)
  x <- inject_slide_title(x = x, title = title, cfg = cfg, ph_label = spec$placeholders$title)
  x <- inject_dotpoints(
    x = x,
    bullets = bullets,
    levels = bullet_levels,
    cfg = cfg,
    ph_label = spec$placeholders$bullets
  )
  x
}
