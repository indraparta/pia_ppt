skip_if_officer_missing <- function() {
  testthat::skip_if_not_installed("officer")
}

skip_if_xml2_missing <- function() {
  testthat::skip_if_not_installed("xml2")
}

template_fixture_path <- function() {
  pkg_file <- system.file(
    "templates",
    "pia_slidemaster.pptx",
    package = "piappt"
  )

  candidates <- c(
    pkg_file,
    file.path(getwd(), "inst", "templates", "pia_slidemaster.pptx"),
    file.path(getwd(), "..", "inst", "templates", "pia_slidemaster.pptx"),
    file.path(getwd(), "..", "..", "inst", "templates", "pia_slidemaster.pptx"),
    file.path(getwd(), "pia_slidemaster.pptx"),
    file.path(getwd(), "..", "pia_slidemaster.pptx"),
    file.path(getwd(), "..", "..", "pia_slidemaster.pptx"),
    "pia_slidemaster.pptx",
    file.path("..", "pia_slidemaster.pptx"),
    file.path("..", "..", "pia_slidemaster.pptx")
  )

  candidates <- unique(candidates[file.exists(candidates)])
  if (!length(candidates)) {
    testthat::skip("Template fixture `pia_slidemaster.pptx` was not found.")
  }

  normalizePath(candidates[[1]], winslash = "/", mustWork = TRUE)
}

first_slide_text <- function(pptx_path) {
  skip_if_xml2_missing()

  td <- tempfile("pptx_unzip_")
  dir.create(td)
  on.exit(unlink(td, recursive = TRUE), add = TRUE)

  utils::unzip(pptx_path, exdir = td)

  pres <- xml2::read_xml(file.path(td, "ppt", "presentation.xml"))
  first_sld <- xml2::xml_find_first(
    pres,
    ".//*[local-name()='sldIdLst']/*[local-name()='sldId'][1]"
  )

  attrs <- xml2::xml_attrs(first_sld)
  rid <- unname(attrs)[length(attrs)]

  rels <- xml2::read_xml(file.path(td, "ppt", "_rels", "presentation.xml.rels"))
  rel_node <- xml2::xml_find_first(
    rels,
    sprintf(".//*[local-name()='Relationship' and @Id='%s']", rid)
  )
  target <- xml2::xml_attr(rel_node, "Target")

  slide <- xml2::read_xml(file.path(td, "ppt", target))
  xml2::xml_text(xml2::xml_find_all(slide, ".//*[local-name()='t']"))
}

first_slide_xml_string <- function(pptx_path) {
  skip_if_xml2_missing()

  td <- tempfile("pptx_unzip_")
  dir.create(td)
  on.exit(unlink(td, recursive = TRUE), add = TRUE)

  utils::unzip(pptx_path, exdir = td)

  pres <- xml2::read_xml(file.path(td, "ppt", "presentation.xml"))
  first_sld <- xml2::xml_find_first(
    pres,
    ".//*[local-name()='sldIdLst']/*[local-name()='sldId'][1]"
  )

  attrs <- xml2::xml_attrs(first_sld)
  rid <- unname(attrs)[length(attrs)]

  rels <- xml2::read_xml(file.path(td, "ppt", "_rels", "presentation.xml.rels"))
  rel_node <- xml2::xml_find_first(
    rels,
    sprintf(".//*[local-name()='Relationship' and @Id='%s']", rid)
  )
  target <- xml2::xml_attr(rel_node, "Target")

  paste(readLines(file.path(td, "ppt", target), warn = FALSE), collapse = "")
}
