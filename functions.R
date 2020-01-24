




loadImage <- function(){
  
  if(file.exists(here("backups/reactiveImage/data.Rdata"))){
    load(here("backups/reactiveImage/data.Rdata"))
  }
  path <- here("backups")
  files <- list.files(path =path, pattern = ".*\\.rds")
  if(length(files)!=0){
    files <- paste0(path, "/",list.files(path =path, pattern = ".*\\.rds"))
    bu <- files[which.max(file.mtime(files))]
    .backup <- readRDS(file = bu)
    .Masterdf <<- .backup
  } else{
    .Masterdf <<- NULL
  }
  #---- load ColIDs
  if(file.exists(here("backups/reactiveImage/data.RData"))){
    load(here("backups/reactiveImage/data.RData"))
    .ColIDs <<- .ColIDs
    .DocPath <<- .DocPath
    .FileNameOutput <<- .FileNameOutput
    .SendTo <<- .SendTo
    .CCoption <<- .CCoption
    .Subject <<- .Subject
    .SendOnBehalf <<- .SendOnBehalf
    .BehalfEmail <<- .BehalfEmail
    .addAttachPath <<- .addAttachPath
    .CanRun <<- .CanRun
  } else { #If backup image isnt saved - set defaults
    .ColIDs <<- NULL
    .DocPath <<-NULL
    .FileNameOutput <<- NULL
    .SendTo <<- NULL
    .CCoption <<- ""
    .Subject <<- "Put your subject here!"
    .SendOnBehalf <<- FALSE
    .BehalfEmail <<- ""
    .addAttachPath <<- NULL
    .CanRun <<- FALSE
  }
  print(paste0("Current Workspace Image:"))
  print(paste0("Output Directory: ",.DocPath))
  print(paste0("Output File Name: ",.FileNameOutput))
  print(paste0("Send To: ", paste0(.SendTo, collapse = ", ")))
  print(paste0("CC: ", .CCoption))
  print(paste0("Subject: ", .Subject))
  print(paste0("Send On Behalf: ", .SendOnBehalf))
  print(paste0("Behalf Email: ", .BehalfEmail))
  print(paste0("Additonal File Path: ", .addAttachPath))
  
}


detectFlags <- function(LongSTR,  startflag = "[", endflag = "]", unique = TRUE) {
  detected <- unlist(str_extract_all(LongSTR, pattern = "\\[[^\\]]+\\]"))
  detected <- gsub(pattern = "\\[|\\]", replacement = "",  detected)
  if(unique){
    detected <- unique(detected)
  }
  return(detected)
}

RDCOMExtractText <- function(docpath){
  docpath <- normalizePath(c("C://Users", docpath), winslash = "\\")[2]
  app <- COMCreate("Word.Application")
  doc <- app[["Documents"]]$Open(docpath, Visible = FALSE)
  docText <- doc$Range()[["Text"]]
  rm(app, doc)
  return(docText)
}




create_latestversion <- function(path, pattern, device = ".csv") {
  x <- list.files(path = path, pattern = pattern)
  if(length(x)==0){
    return(paste0(path,"\\",pattern,device))
  }
  x2 <- str_split(string = x, pattern = paste0(pattern,"|\\",device))
  x <- as.numeric(unlist(lapply(x2, FUN = function(x){x[2]})))
  x[is.na(x)] <- 0
  index <- max(x)+1
  newpath <- paste0(path,"\\",pattern,index,device)
  return(newpath)
}





FixDTCoerce <- function(data){
  l1 <- lapply(data, grep, pattern = "&amp;", fixed =T)
  for(n in names(l1)){
    if(length(l1[[n]])>0){
      data[[n]] <- gsub(pattern = "&amp;", replacement = "&", x = data[[n]], fixed = T)
    }
  }
  return(data)
}

flagInColumn <- function(str, flag){
  t <- str[unlist(lapply(str, grepl, flag))]
  if(length(t)==0){
    return(NA)
  } else{
    return(t)
  }
}

sub_flags <- function(x, data){
  n <- length(x)
  x <- gsub(pattern = "\\[|\\]", replacement = "", x = x)
  cnames <- colnames(data)
  for(i in 1:n){
    if(x[i]%in%cnames){
      x[i] <- as.character(data[[as.character(x[i])]])
    }
  }
  return(x)
}


sub_flag_in_str <- function(x, flags, data) {
  avail <-  flags %in% colnames(data)
  if(length(flags)!=sum(avail)){
    flags <- flags[avail]
  }
  for(f in flags){
    if(is.na(data[[f]])){
      r <- ""
    } else {
      r <- as.character(data[[f]])
    }
    x <- gsub(pattern = paste0("[",f,"]"), replacement = r, x = x, fixed = TRUE, )
  }
  return(x)
}


#currently not used but may be useful??
make_str_literal <- function(string,
                             escape.characters = c("!","+","[","]","(",")","*","."," ","^","&", "$","@", "{","}","|",",")){
  #print(string)
  positions <- unlist(gregexpr(pattern = paste0("\\",escape.characters, collapse = "|"), string))
  if(length(positions)==1&&positions==-1){
    return(string)
  }
  for(i in length(positions):1){
    string <- paste0(substr(string,0,(positions[i]-1)),"\\",str_sub(string, start = positions[i]))
  }
  return(string)
}

#testing
# doc <- read_docx("../../testdoc.docx")
# nodes <- getnodes(doc)


test_wordapp <- function(app){
  tryCatch({app[["Documents"]]
    return(app)},
    error=function(cond){
      print("Word app does not seem to be accessible, Returning a new app")
      app <- COMCreate("Word.Application")
      return(app)
    })
}

RDCOMFindReplace <- function(flags,
                             data,
                             docpath,
                             targetdir,
                             wordApp,
                             makeWhich = "both",
                             pw = NULL){
  extensions <- c()
  er <- NULL
  wordApp <- test_wordapp(wordApp)
  docpath <- normalizePath(c("C://Users", docpath), winslash = "\\")[2]
  suppressWarnings(targetdir <- normalizePath(c("C://Users", targetdir), winslash = "\\")[2])
  doc <- wordApp[["Documents"]]$Open(docpath, Visible = FALSE)
  
  for(f in flags){
    replace <- data[[f]]
    replace <- ifelse(is.na(replace), "",as.character(replace))
    doc$Range()[["Find"]]$Execute(FindText = paste0("[",f,"]"),
                                  ReplaceWith = replace, Replace = 2) #Uses RDCOMClient to replace fields.
    targetdir <- gsub(paste0("[",f,"]"), gsub(pattern = "\\/|\"",
                                              replacement = "_",
                                              x = replace), x = targetdir, fixed = TRUE)
  }
  docText <- doc$Range()[["Text"]]
  if(length(detectFlags(docText))>0){
    print("Flag Replacement Error")
    er <- "Document Flag Replacement"
  }
  
  if(makeWhich %in% c("both","docx")){
    if(!is.null(pw)&&str_length(pw)>0){
      doc$Protect(0, password = pw)
    }
    doc$SaveAs(paste0(targetdir,".docx"))
    extensions <- c(extensions, ".docx")
  }
  if(makeWhich %in% c("both","pdf")){
    doc$SaveAs(paste0(targetdir,".pdf"), FileFormat=17) #saves as PDF
    extensions <- c(extensions, ".pdf")
  }
  doc$Close(SaveChanges = 0)
  rm(doc)
  retvalue <- list(str=str_flatten(paste0(targetdir,extensions), "|.|"), error = er)
  return(retvalue)

  
}


RDCOMpassword <- function(docpath, pw){
  if(file.exists(docpath)){
    docpath <- normalizePath(c("C://Users", docpath), winslash = "\\")[2]
    app <- COMCreate("Word.Application")
    doc <- app[["Documents"]]$Open(docpath, Visible = FALSE)
    doc[["Password"]] <- pw #sets password on document
   # doc$Protect(0, password = pw) #Sets change
    doc$Close(SaveChanges = -1)
    rm(app, doc)
  }
}



