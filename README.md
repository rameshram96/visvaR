# *visvaR - Visualize Variance* <img src="man/figures/visvaRlogo.png" align="right" height="200" style="float:right; height:200px;"/>

It is a user-friendly tool to do the statistical analysis of agricultural research data like ANOVA, correlation, etc. (This package is in the developmental stage, we strongly encourage feedback from the users to improve the package) Package developed by Ramesh R PhD Scholar, Division of Plant Physiology ICAR-IARI, New Delhi [ramesh.rahu96\@gmail.com](mailto:ramesh.rahu96@gmail.com){.email}

## *Installation*

You can install the development version of visvaR from [GitHub](https://github.com/rameshram96/visvaR) using the following command:

``` r
install.packages("devtools")
devtools::install_github("rameshram96/visvaR")
```

## *Example usage*

``` r
library(visvaR)
visvaR:::oneway_crd()        # one factor completely randomized design 
visvaR:::oneway_rbd()        # one factor randomized block design 
visvaR:::twoway_crd()        # two factor completely randomized design 
visvaR:::twoway_rbd()        # two factor randomized block design 
visvaR:::correlation_multi() # Correlation of multiple variables 
```

Running any of this code will open a application in your browser, where you can paste or import your data, after analysing the data you can download the results as a word file.

## *Using custom fonts inside the pacakge*

Make sure that system fonts are imported properly and loaded

``` r
install.packages("extrafont")
library(extrafont)
extrafont::font_import()
extrafont::loadfonts()
```

<!-- badges: start -->

[![R-CMD-check](https://github.com/rameshram96/visvaR/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/rameshram96/visvaR/actions/workflows/R-CMD-check.yaml) [![CRAN status](https://www.r-pkg.org/badges/version/visvaR)](https://CRAN.R-project.org/package=visvaR)

<!-- badges: end -->
