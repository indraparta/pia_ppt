test_that("new_pia_template builds expected content and front specs", {
  path <- template_fixture_path()
  cfg <- new_pia_template(path)

  expect_s3_class(cfg, "piappt_template")
  expect_equal(cfg$content_spec$layout, "03_Text-Chart-2")
  expect_equal(cfg$front_spec$layout, "01_Title")
  expect_equal(cfg$content_spec$placeholders$chart, "Content Placeholder 12")
  expect_equal(cfg$front_spec$placeholders$subtitle, "Subtitle 2")
})

test_that("new_pia_template can use bundled template by default", {
  cfg <- new_pia_template()
  expect_true(file.exists(cfg$template_path))
  expect_equal(basename(cfg$template_path), "pia_slidemaster.pptx")
})

test_that("new_pia_template enforces required front placeholders", {
  path <- template_fixture_path()

  expect_error(
    new_pia_template(
      template_path = path,
      front_placeholders = list(title = "Title 1")
    ),
    "front_placeholders.*subtitle"
  )
})

test_that("open_template defaults to new-slides-only mode", {
  skip_if_officer_missing()

  cfg <- new_pia_template(template_fixture_path())
  x <- open_template(cfg)

  # Default behavior should strip existing template slides.
  expect_equal(length(x), 0)
})
