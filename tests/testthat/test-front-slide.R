test_that("add_front_slide inserts a front slide and places it at index 1", {
  skip_if_officer_missing()
  skip_if_xml2_missing()

  cfg <- new_pia_template(template_fixture_path())
  x <- open_template(cfg)
  x <- add_master_slide(x = x, cfg = cfg)

  front_title <- paste0("Front Title ", as.integer(Sys.time()))
  x <- add_front_slide(x = x, cfg = cfg, title = front_title)

  out <- tempfile(fileext = ".pptx")
  write_presentation(x, out)
  first_text <- first_slide_text(out)
  expect_true(front_title %in% first_text)
})

test_that("add_front_slide injects subtitle when provided", {
  skip_if_officer_missing()
  skip_if_xml2_missing()

  cfg <- new_pia_template(template_fixture_path())
  x <- open_template(cfg)

  front_title <- paste0("Front Title ", as.integer(Sys.time()))
  front_subtitle <- paste0("Front Subtitle ", as.integer(Sys.time()))
  x <- add_front_slide(x = x, cfg = cfg, title = front_title, subtitle = front_subtitle)

  out <- tempfile(fileext = ".pptx")
  write_presentation(x, out)
  first_text <- first_slide_text(out)
  expect_true(front_subtitle %in% first_text)
})

test_that("add_front_slide requires a non-empty title", {
  skip_if_officer_missing()

  cfg <- new_pia_template(template_fixture_path())
  x <- open_template(cfg)

  expect_error(
    add_front_slide(x = x, cfg = cfg, title = ""),
    "`title` must be a non-empty string."
  )
})

test_that("add_front_slide reports available labels when placeholder is missing", {
  skip_if_officer_missing()

  cfg <- new_pia_template(
    template_path = template_fixture_path(),
    front_placeholders = list(
      title = "Title 1",
      subtitle = "Not A Real Placeholder"
    )
  )
  x <- open_template(cfg)

  expect_error(
    add_front_slide(x = x, cfg = cfg, title = "Front title"),
    "Available labels"
  )
})
