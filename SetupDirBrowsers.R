.posRoot <- paste(str_split(getwd(),pattern = "/")[[1]][1:3],collapse = "/")
.files <- list.files(.posRoot)
.files <- .files[.files %in% c("Desktop","Documents","Downloads")]
.p <- paste0(.posRoot,"/",c("Desktop","Documents","Downloads"))
.pavail <- c("Desktop","Documents","Downloads") %in% .files
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
