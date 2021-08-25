# viewxl

> An addin to interactively export `data.frame` from global environment into Excel with [`writexl`](https://github.com/ropensci/writexl).


<!-- badges: start -->
[![Project Status: Active – The project has reached a stable, usable state and is being actively developed.](http://www.repostatus.org/badges/latest/active.svg)](http://www.repostatus.org/#active)
[![R-CMD-check](https://github.com/dreamRs/viewxl/workflows/R-CMD-check/badge.svg)](https://github.com/dreamRs/viewxl/actions)
<!-- badges: end -->


## Installation

Install from GitHub with:

```r
remotes::install_github("dreamRs/viewxl")
```


## Usage


* Select a `data.frame` then launch the addin via RStudio addins menu or with shortcut, the selected `data.frame` will be open in Excel directly:

![](screenshots/selection_df.png)


* If no selection is made, a Shiny gadget is launched to select `data.frame` to open:

![](screenshots/addin.png)


* Or use in console with:

```r
viewxl::view_in_xl(mtcars)
``` 
