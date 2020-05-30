# The purpose of this script is to generate example output plots to population
# the report.  We will be creating a series of random walk curves. Then, create
# an automated powerpoint presentation report using the officer package.

# SETUP -------------------------------------------------------------------

library(tidyverse)
library(lubridate)
library(here)
library(glue)
library(officer)

# Define parameters
n_scenarios <- 5
mean_min <- -2
mean_max <- 2
sd_min   <- 10
sd_max   <- 25

# DEFINE FUNCTION FOR RANDOM WALK -----------------------------------------

generate_random_walk <- function(mean = 0, 
                                 sd = 1,
                                 n = 366,
                                 x0 = ymd("2020-01-01"), 
                                 y0 = 1000) {
  
  data <- tibble(
    x = x0 + 1:n,
    y = y0 + rnorm(n, mean, sd) %>% cumsum()
  )
  
  return(data)
}


# GENERATE DATA -----------------------------------------------------------

# Define list of scenarios
scenarios = str_to_upper(letters)[1:n_scenarios]

# Initialize grid of scenarios
df_randwalk <- 
  tibble(
    scenario = scenarios,
    mean = runif(n_scenarios, mean_min, mean_max),
    sd = runif(n_scenarios, sd_min, sd_max)
  )

# Generate random walk data
df_randwalk <- df_randwalk %>%
  mutate(data = map2(mean, sd, generate_random_walk)) %>%
  unnest(data)

# CREATE PLOTS ------------------------------------------------------------

for (scenario_target in scenarios) {
  
  
  # Gather target scenario
  df_target <- df_randwalk %>% filter(scenario == scenario_target)
  
  # Gather other scenarios
  df_others <- df_randwalk %>% filter(scenario != scenario_target)
  
  # Create plot
  df_randwalk %>%
    ggplot() + 
    geom_line(
      data = df_others, 
      mapping = aes(x = x, y = y, group = scenario),
      alpha = 0.25, color = "grey50"
    ) +
    geom_line(
      data = df_target, 
      mapping = aes(x = x, y = y), 
      color = "dodgerblue3", size = 1
    ) +
    scale_y_continuous(labels = scales::dollar) +
    theme_light() +
    labs(
      title = "Comparitive Analysis",
      subtitle = glue("Scenario {scenario_target}"),
      x = "Date",
      y = "Sales",
      color = "Scenario"
    ) +
    ggsave(
      here("plots", glue("plot_scenario_{scenario_target}.jpg")),
      height = 5, width = 11, dpi = 1000
    )
  
}


# CREATE REPORT -----------------------------------------------------------

# Initialize report
report <- read_pptx(here("report", "template.pptx"))

# Add slides
for (scenario_target in scenarios) {
  
  # Add slide
  report <- report %>%
    add_slide(layout = "Title and Content", master = "RetrospectVTI") %>%
    ph_with(
      value = glue("Scenario {scenario_target}"), 
      location = ph_location_type(type = "title")
    ) %>%
    ph_with(
      value = external_img(here("plots", glue("plot_scenario_{scenario_target}.jpg")), width = 11, height = 5),
      location = ph_location_type(type = "body"), use_loc_size = TRUE
    )
  
}

# Output report
print(report, here("report", "report.pptx"))

