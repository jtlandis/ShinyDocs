.posRoot <- paste(str_split(getwd(),pattern = "/")[[1]][1:3],collapse = "/")
.files <- list.files(.posRoot)
.locations <- c("Desktop","Documents","Downloads")
.files <- .files[.files %in% .locations]
.p <- paste0(.posRoot,"/",.locations)
.pavail <- .locations %in% .files
if(length(.pavail)==0){
  dirList <- list(home = "~", rootR = here())
} else {
  .p <- .p[order(.p)]
  dirList <- list(Desktop = .p[1],
                  Documents = .p[2],
                  Downloads = .p[3],
                  rootR = here())
  dirList[c(.pavail,TRUE)]
}
