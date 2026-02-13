# piappt

Lightweight helpers for generating PowerPoint decks from an existing slide
master template.

## Workflow

```r
library(piappt)
library(ggplot2)

cfg <- new_pia_template() # uses bundled template in inst/templates
# Equivalent explicit path:
# template_path <- system.file("templates", "pia_slidemaster.pptx", package = "piappt")
# cfg <- new_pia_template(template_path)

ppt <- open_template(cfg) # default: strip existing template slides

# Front page (01_Title)
ppt <- add_front_slide(
  x = ppt,
  cfg = cfg,
  title = "Project Scoping Pack",
  subtitle = "Policy Institute"
)

# Primary content slide (03_Text-Chart-2)
chart <- ggplot(mtcars, aes(wt, mpg)) + geom_point()

ppt <- add_primary_slide(
  x = ppt,
  cfg = cfg,
  chart = chart,
  title = "Fuel efficiency overview",
  bullets = c(
    "Heavier cars generally have lower MPG.",
    "Most observations cluster around 2 to 4 tons.",
    "Placement follows named slide placeholders."
  )
)

# Bullet paragraph spacing defaults to 3pt before and 3pt after.
# Override via:
# bullet_spacing_before_pt = 5, bullet_spacing_after_pt = 7

write_presentation(ppt, "output/deck.pptx")
```

Use `list_placeholders(cfg$template_path, layout = "03_Text-Chart-2")` if you
want to verify or customize placeholder labels.

If you want to keep all original template slides, use:

```r
ppt <- open_template(cfg, include_existing_slides = TRUE)
```
