assert_single_string <- function(x, arg_name) {
  if (!is.character(x) || length(x) != 1 || is.na(x) || !nzchar(trimws(x))) {
    stop("`", arg_name, "` must be a non-empty string.", call. = FALSE)
  }
}

assert_optional_single_string <- function(x, arg_name) {
  if (is.null(x)) {
    return(invisible(NULL))
  }

  assert_single_string(x, arg_name)
}

assert_single_logical <- function(x, arg_name) {
  if (!is.logical(x) || length(x) != 1 || is.na(x)) {
    stop("`", arg_name, "` must be TRUE or FALSE.", call. = FALSE)
  }
}

assert_rpptx <- function(x, arg_name) {
  if (!inherits(x, "rpptx")) {
    stop("`", arg_name, "` must be an `officer::rpptx` object.", call. = FALSE)
  }
}

assert_template_config <- function(cfg) {
  if (!inherits(cfg, "piappt_template")) {
    stop("`cfg` must be created with `new_pia_template()`.", call. = FALSE)
  }

  required <- c("template_path", "content_spec", "front_spec")
  missing <- setdiff(required, names(cfg))
  if (length(missing) > 0) {
    stop(
      "`cfg` is missing required fields: ",
      paste(missing, collapse = ", "),
      call. = FALSE
    )
  }

  if (!file.exists(cfg$template_path)) {
    stop("Template file does not exist: ", cfg$template_path, call. = FALSE)
  }
}

validate_placeholder_map <- function(map, required_names, arg_name) {
  if (!is.list(map) || is.null(names(map))) {
    stop("`", arg_name, "` must be a named list.", call. = FALSE)
  }

  missing <- setdiff(required_names, names(map))
  if (length(missing) > 0) {
    stop(
      "`", arg_name, "` is missing required keys: ",
      paste(missing, collapse = ", "),
      call. = FALSE
    )
  }

  out <- map[required_names]
  for (name in required_names) {
    assert_single_string(out[[name]], paste0(arg_name, "$", name))
  }

  out
}

resolve_layout_master <- function(x, layout, master = NULL) {
  assert_rpptx(x, "x")
  assert_single_string(layout, "layout")
  assert_optional_single_string(master, "master")

  summary <- officer::layout_summary(x)
  if (!all(c("layout", "master") %in% names(summary))) {
    stop("`officer::layout_summary()` did not return expected columns.", call. = FALSE)
  }

  rows <- summary[summary$layout == layout, , drop = FALSE]
  if (nrow(rows) == 0) {
    stop(
      "Layout `", layout, "` was not found. Available layouts: ",
      paste(unique(summary$layout), collapse = ", "),
      call. = FALSE
    )
  }

  if (is.null(master)) {
    master <- rows$master[[1]]
  } else if (!master %in% rows$master) {
    stop(
      "Layout `", layout, "` is not available under master `", master, "`. ",
      "Available masters for this layout: ",
      paste(unique(rows$master), collapse = ", "),
      call. = FALSE
    )
  }

  list(layout = layout, master = master)
}

placeholder_labels <- function(x, layout, master) {
  props <- officer::layout_properties(x, layout = layout, master = master)
  labels <- unique(props$ph_label)
  labels <- labels[!is.na(labels) & nzchar(labels)]
  labels
}

assert_placeholders_exist <- function(x, layout, master, labels, context) {
  available <- placeholder_labels(x, layout = layout, master = master)
  missing <- setdiff(labels, available)

  if (length(missing) > 0) {
    stop(
      "Missing placeholder label(s) for ", context, ": ",
      paste(missing, collapse = ", "),
      ". Available labels for layout `", layout, "` and master `", master, "`: ",
      paste(available, collapse = ", "),
      call. = FALSE
    )
  }
}

resolve_location <- function(ph_label = NULL, location = NULL) {
  if (!is.null(location)) {
    return(location)
  }

  assert_single_string(ph_label, "ph_label")
  officer::ph_location_label(ph_label = ph_label)
}

slide_count <- function(x) {
  assert_rpptx(x, "x")

  n <- tryCatch(
    as.integer(length(x)),
    error = function(...) NA_integer_
  )

  if (!is.na(n) && n >= 0) {
    return(n)
  }

  if (!exists("pptx_summary", envir = asNamespace("officer"), inherits = FALSE)) {
    stop("Unable to determine slide count for `x`.", call. = FALSE)
  }

  pptx_summary <- getFromNamespace("pptx_summary", "officer")
  s <- pptx_summary(x)
  if (!nrow(s) || !"slide_id" %in% names(s)) {
    return(0L)
  }

  as.integer(max(s$slide_id, na.rm = TRUE))
}

move_slide_to_start <- function(x, from_index) {
  assert_rpptx(x, "x")

  if (!exists("move_slide", envir = asNamespace("officer"), inherits = FALSE)) {
    stop(
      "Front slide was created, but move-to-start failed because `officer::move_slide()` ",
      "is not available. Recommended remediation: call `add_front_slide()` before adding ",
      "other slides, or update `officer`.",
      call. = FALSE
    )
  }

  move_slide <- getFromNamespace("move_slide", "officer")
  args <- names(formals(move_slide))

  call_args <- if (all(c("x", "index", "to") %in% args)) {
    list(x = x, index = from_index, to = 1L)
  } else if (all(c("x", "from", "to") %in% args)) {
    list(x = x, from = from_index, to = 1L)
  } else if (all(c("x", "slide_index", "to") %in% args)) {
    list(x = x, slide_index = from_index, to = 1L)
  } else {
    stop(
      "Front slide was created, but move-to-start failed due to an unsupported ",
      "`officer::move_slide()` signature. Recommended remediation: call ",
      "`add_front_slide()` before adding other slides, or update `officer`.",
      call. = FALSE
    )
  }

  tryCatch(
    do.call(move_slide, call_args),
    error = function(e) {
      stop(
        "Front slide was created, but move-to-start failed. Recommended remediation: ",
        "call `add_front_slide()` before adding other slides, or update `officer`. ",
        "Underlying error: ", conditionMessage(e),
        call. = FALSE
      )
    }
  )
}

remove_slide_at <- function(x, index) {
  assert_rpptx(x, "x")

  if (!exists("remove_slide", envir = asNamespace("officer"), inherits = FALSE)) {
    stop(
      "Cannot remove existing slides because `officer::remove_slide()` is unavailable. ",
      "Open with `include_existing_slides = TRUE` or update `officer`.",
      call. = FALSE
    )
  }

  remove_slide <- getFromNamespace("remove_slide", "officer")
  args <- names(formals(remove_slide))

  call_args <- if (all(c("x", "index") %in% args)) {
    list(x = x, index = as.integer(index))
  } else if (all(c("x", "from") %in% args)) {
    list(x = x, from = as.integer(index))
  } else if (all(c("x", "slide_index") %in% args)) {
    list(x = x, slide_index = as.integer(index))
  } else {
    stop(
      "Cannot remove existing slides due to an unsupported `officer::remove_slide()` ",
      "signature. Open with `include_existing_slides = TRUE` or update `officer`.",
      call. = FALSE
    )
  }

  do.call(remove_slide, call_args)
}

strip_all_slides <- function(x) {
  assert_rpptx(x, "x")

  n <- slide_count(x)
  if (n <= 0L) {
    return(x)
  }

  for (i in seq_len(n)) {
    x <- remove_slide_at(x, index = 1L)
  }

  x
}
