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
#' @param spacing_before_pt Paragraph spacing before each bullet, in points.
#'   Defaults to `3`.
#' @param spacing_after_pt Paragraph spacing after each bullet, in points.
#'   Defaults to `3`.
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
  spacing_before_pt = 3,
  spacing_after_pt = 3,
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
  if (!is.numeric(spacing_before_pt) || length(spacing_before_pt) != 1 ||
      is.na(spacing_before_pt) || spacing_before_pt < 0) {
    stop("`spacing_before_pt` must be a non-negative number.", call. = FALSE)
  }
  if (!is.numeric(spacing_after_pt) || length(spacing_after_pt) != 1 ||
      is.na(spacing_after_pt) || spacing_after_pt < 0) {
    stop("`spacing_after_pt` must be a non-negative number.", call. = FALSE)
  }

  loc <- resolve_location(ph_label = ph_label, location = location)
  ph_with_spaced_unordered_list(
    x = x,
    bullets = bullets,
    levels = as.integer(levels),
    location = loc,
    spacing_before_pt = spacing_before_pt,
    spacing_after_pt = spacing_after_pt
  )
}

#' Add a primary content slide
#'
#' @param x An `officer::rpptx` object.
#' @param cfg A configuration object created by `new_pia_template()`.
#' @param chart A chart-like object supported by `officer::ph_with()`.
#' @param title A title string for the content slide.
#' @param bullets Character vector of bullet text.
#' @param bullet_levels Optional integer vector of bullet indentation levels.
#' @param bullet_spacing_before_pt Paragraph spacing before each bullet, in
#'   points. Defaults to `3`.
#' @param bullet_spacing_after_pt Paragraph spacing after each bullet, in
#'   points. Defaults to `3`.
#'
#' @return Updated `officer::rpptx` object.
#' @export
add_primary_slide <- function(
  x,
  cfg,
  chart,
  title,
  bullets,
  bullet_levels = NULL,
  bullet_spacing_before_pt = 3,
  bullet_spacing_after_pt = 3
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
    spacing_before_pt = bullet_spacing_before_pt,
    spacing_after_pt = bullet_spacing_after_pt,
    cfg = cfg,
    ph_label = spec$placeholders$bullets
  )
  x
}

ph_with_spaced_unordered_list <- function(
  x,
  bullets,
  levels,
  location,
  spacing_before_pt = 3,
  spacing_after_pt = 3
) {
  slide <- x$slide$get_slide(x$cursor)
  fortify_location <- getFromNamespace("fortify_location", "officer")
  shape_properties_tags <- getFromNamespace("shape_properties_tags", "officer")
  html_escape <- getFromNamespace("htmlEscapeCopy", "officer")
  psp_ns_yes <- getFromNamespace("psp_ns_yes", "officer")

  location <- fortify_location(location, doc = x)
  new_ph <- shape_properties_tags(
    left = location$left,
    top = location$top,
    width = location$width,
    height = location$height,
    label = location$ph_label,
    ph = location$ph,
    rot = location$rotation,
    bg = location$bg,
    ln = location$ln,
    geom = location$geom
  )

  before_val <- as.integer(round(spacing_before_pt * 100))
  after_val <- as.integer(round(spacing_after_pt * 100))

  pars <- vapply(seq_along(bullets), function(i) {
    lvl <- as.integer(levels[[i]])
    lvl_attr <- if (lvl > 1L) sprintf(" lvl=\"%d\"", lvl - 1L) else ""
    bu_none <- if (lvl < 1L) "<a:buNone/>" else ""

    sprintf(
      paste0(
        "<a:p><a:pPr%s>",
        "<a:spcBef><a:spcPts val=\"%d\"/></a:spcBef>",
        "<a:spcAft><a:spcPts val=\"%d\"/></a:spcAft>",
        "%s</a:pPr><a:r><a:rPr/><a:t>%s</a:t></a:r></a:p>"
      ),
      lvl_attr,
      before_val,
      after_val,
      bu_none,
      html_escape(as.character(bullets[[i]]))
    )
  }, character(1))

  xml_elt <- paste0(
    psp_ns_yes,
    new_ph,
    "<p:txBody><a:bodyPr/><a:lstStyle/>",
    paste0(pars, collapse = ""),
    "</p:txBody></p:sp>"
  )
  node <- xml2::as_xml_document(xml_elt)
  xml2::xml_add_child(xml2::xml_find_first(slide$get(), "//p:spTree"), node)
  x
}
