# devtools::install_github("omegahat/RDCOMClient")
library(RDCOMClient)
library(shiny)
library(shinyFiles)
library(DT)
library(readxl)
library(stringr)
library(shinyMCE)
library(ggplot2)
library(here)

# if(!require("devtools")){
# install.packages("devtools") 
# devtools::install_github("mul118/shinyMCE")
# }
# loadMCEcontent <- function(){
#   if(file.exists(here("backups/reactiveImage/MCEcontent.rds"))){
#     x <- readRDS(here("backups/reactiveImage/MCEcontent.rds"))
#   } else {
#     x <- "Write Your Email Here! Please use \"[ ]\" to idicate replacement fields."
#   }
#   return(x)
# }
# saveMCEcontent <- function(x){
#   if(!dir.exists(here("backups/reactiveImage/"))){
#     dir.create(here("backups/reactiveImage/"))
#   }
#   saveRDS(object = x, file = here("backups/reactiveImage/MCEcontent.rds"))
# }

ui <- shinyUI(
  fluidPage(
    tags$head(
      includeCSS('www/style.css'),
      tags$link(
        rel = "icon",
        type = "image/x-icon",
        href = "http://localhost:1984/EZDocs.ico"
      )
    ),
    h1(strong("EZDocs"), align = "center", style = "color:#FF6600;font-family: 'Lobster', cursive;"),
    #actionButton("browser","browser"),
    tabsetPanel(
      tabPanel("Document Setup",
               wellPanel(
                 fluidRow(
                   column(6,
                          fileInput("ChooseTemplate",
                                    "Upload Word Template File (.docx)",
                                    accept = c(".docx")
                          ),
                          htmlOutput("DetectedIDs"),
                          fluidRow(
                            column(6,
                                   br(),
                                   shinyDirButton("dir",
                                                  "Choose Output Folder",
                                                  title = "Upload"),
                                   htmlOutput(outputId = "dirDisplay")
                            ),
                            column(6,
                                   uiOutput("DynamicInput"),
                                   htmlOutput("MessageFileNameOutputDynam")
                            )
                          )
                   ),
                   column(6,
                          tabsetPanel(
                            tabPanel("Excel Import",
                                     column(12,
                                            br(),
                                            fileInput("TableIn",
                                                      "Upload Table (Excel, .CSV format)",
                                                      accept = c(".xls",".xlsx",".csv")
                                            ),
                                            uiOutput(outputId = "SheetSelection")
                                            
                                     )
                            ),
                            tabPanel("Backup Import",
                                     column(12, align = "center",
                                            br(),
                                            "Load from Backup?",
                                            br(),
                                            actionButton("backupLoad", "Load Backup")
                                     )
                            )
                          )
                   )
                 ),
                 br(),
                 fluidRow(
                   column(3, offset = 3,
                          align = "center",
                          actionButton("MakeDocuments",
                                       "Make Documents")),
                   column(3,
                          align = "center",
                          actionButton("ExportAllData",
                                       "Export All")
                   )
                 )
               )
      ),
      tabPanel("Email",
               wellPanel(
                 fluidRow(
                   column(6,
                          tinyMCE("EmailEditor",
                                  content = "Write Your Email Here! Please use \"[ ]\" to idicate replacement fields"),
                          htmlOutput("EmailOutputsignature")
                   ),
                   uiOutput("EmailPanel")
                 ),
                 
                 br(),
                 fluidRow(
                   column(12,
                          uiOutput("NAflags")
                   )
                 ),
                 br(),
                 fluidRow(
                   column(12,
                          align = "center",
                          actionButton("SendEmails",
                                       "Send Emails")
                   )
                 )
               )
      ),
      tabPanel("Help", #Configure Back up number, How to handle NAs in documents. In Emails
               #          uiOutput("OptionsPanel"),
               wellPanel(
                 h2(strong("Welcome to EZDocs!"), align = "center", style = "color:#FF6600;font-family: 'Lobster', cursive;"),
                 br(),
                 fluidRow(
                   column(2, class="sticky",
                          tags$ul(class="nav nav-pills nav-stacked shiny-tab-input",
                                  tags$li(class = "active", tags$a(href="#tab-5000-1",
                                                                   'data-toggle'="tab",
                                                                   'data-value'="Goal",
                                                                   "Goal")
                                          ),
                                  tags$li(
                                    tags$a(href="#tab-5000-2",
                                           'data-toggle'="tab",
                                           'data-value'="Requirements",
                                           "Requirements")
                                  ),
                                  tags$li(
                                    tags$a(href="#tab-5000-3",
                                           'data-toggle'="tab",
                                           'data-value'="Flags",
                                           "Flags")
                                  ),
                                  tags$li(
                                    tags$a(href="#tab-5000-4",
                                           'data-toggle'="tab",
                                           'data-value'="Descriptions",
                                           "Descriptions")
                                  ),
                                  tags$li(
                                    tags$a(href="#tab-5000-5",
                                           'data-toggle'="tab",
                                           'data-value'="Data Table",
                                           'Data Table')
                                  ),
                                  tags$li(
                                    tags$a(href="#tab-5000-6",
                                           'data-toggle'="tab",
                                           'data-value'="Backups",
                                           "Backups")
                                  ),
                                  tags$li(
                                    tags$a(href="#tab-5000-7",
                                           'data-toggle'="tab",
                                           'data-value'="Errors",
                                           "Errors")
                                  )
                              )
                          ),
                   column(10, #Tab Content...
                          div(class="tab-content", 'data-tabsetid'="5000",
                              div(class="tab-pane active",
                                  'data-value'="Goal",
                                  id="tab-5000-1",
                                  h3(strong("Goal of Application")),
                                  p("Create multiple Word and Adobe PDF documents with ease while allowing for the user
                   to send documents efficiently and with minimal errors. For each row in the Data Frame,
                   any flags and with the exact name as the column in [brackets], found within the document
                   will be substituted with the value in the cell on that column."),
                                
                              ),
                              div(class="tab-pane",
                                  'data-value'="Requirements",
                                  id="tab-5000-2",
                                  h3(strong("Expected applications")),
                                  tags$ul(
                                    tags$li("Working Outlook application installed on computer"),
                                    tags$li("Template Document (.docx) with \"[flags]\""),
                                    tags$li("Excel Document to substitute values into \"[flags]\""),
                                    tags$li("Windows Computer")
                                    )
                                  ),
                              div(class="tab-pane",
                                  'data-value'="Flags",
                                  id="tab-5000-3",
                                  h3(strong("Flags ---- An Explanation")),
                                  p("It is important to know what flags mean in the EZDocs app. EZDocs
             operates around Flags, which is anything contained between two brackets, for example \"[bracket]\". Flags 
             may be used anywhere
             in EZDocs (Document name, Email header, Customized Email, Inside the document requiring customization, etc.) 
             Only text that exactly match the column headers of
             the imported Data Frame are valid Flags. In most cases the application
             will attempt to warn you if it detects a flag that is invalid. Flags should be used in at least 
             the template.docx file and the Customize New File Name field. Flags may be used in the email section as
             well.")
                                  ),
                              div(class="tab-pane",
                                  'data-value'="Descriptions",
                                  id="tab-5000-4",
                                  h3(strong("How to use ---- Features")),
                                  tags$ul(
                                    tags$li(h4(strong("Document Setup")),
                                            tags$ul(
                                              tags$li(strong("Upload Word Template File"),
                                                      p("Opens a file browser to select the Word document template. Restrictions in the code requires that the uploaded word document
		is a .docx file. Once a file is chosen the application will notify the user of how many flags it detects and
		how many of those flags are valid (Also found within the excel spreadsheet). If any flags are invalid it will notify the user. The user will
		not be able to make documents until all flags are valid")
                                              ),
                                              tags$li(strong("Choose Output Folder"),
                                                      p("Opens a window to select an output directory (where documents will be stored). By default four locations are chosen: Desktop, Documents, Downloads, and rootR.
		The rootR directory is where the application will be running from and is where it saves its backup Rdata files and R code. It is not recommended
		for the user to use this directory to store files and is only listed such that the user may find it if needed. Once a directory is chosen
		it will output the path under the button. Users may copy this path into their Desktop's file explorer to navigate to the location of the documents quickly.")
                                              ),
                                              tags$li(strong("Customize New File Name"),
                                                      p("This field specifies how each document will be named. It is necessary that the user uses a combination of flags, so
                                     the expected output names are unique. This is to prevent the user from overwriting other documents automatically.
		This does not check for any NA values. Documents whose names contain an NA values are not written. This is because the 
		application does not know what to substitute for the flag whose value is NA. This is not the case with empty strings \"\" (no space). Any
		document that is attempted to be made with an NA value in its name will output an Error in the error log file of the 
		specified Output Directory. The specific error it will throw is \"Flag Value\". See the Error section for more Details.")
                                              ),
                                              tags$li(strong("Excel Import"),
                                                      p("Import an excel file and select which sheet to use. Once a data frame is imported it will be visible at the
		bottom of the application. Please see the \"Data table Features\" sections to read more.")
                                              ),
                                              tags$li(strong("Backup Import"),
                                                      p("Load from a previous version of the Data Frames used. These files are stored as .rds files in the 
		Documents\\EZDocs\\backups directory. Any Time the imported Data Frame is modified, a new backup is stored. Current default is set to 12 backups.")
                                              ),
                                              tags$li(strong("Make Documents"),
                                                      p("Creates a word .docx and .pdf files with the specified output names and proper flag replacement values. Any errors that
		occur during this step is recorded in the error log file in the same output directory. This output is for systematic error which are different from random errors.
		See the errors section for more details. The file path for each
                                       row is recorded in the \"DocumentPaths\" column and the \"DocumentsMade\" records the name of the respective file. A user
                                       may always add additional file paths in this column but must make sure to separate file paths with the special 
                                       characters \"|.|\". The sequence \"|.|\" was chosen in order to indicate when multiple files are entered, the program will treat them as separate strings."),
                                                      strong("NOTE:"), 
                                                      p(" the user can press this button more than once. The program will overwrite any file the file that already exists in
                                     the output directory. If the file name is different, then the application will append the new file path into the proper cell
                                     without modifying existing data.")
                                              ),
                                              tags$li(strong("Export All"),
                                                      p("Exports the entire Data Frame to a .csv file. The default path is to the currently set Output Directory.")
                                              )
                                            )
                                    ),
                                    
                                    tags$li(h4(strong("Emails")),
                                            tags$ul(
                                              tags$li(strong("Email Editor"),
                                                      p("Use the text Box to write the email. Flags may be used here as well to send emails dynamically. When using flags, 
                                       be sure that any formatting is on the entire flag, including the brackets.")),
                                              tags$li(strong("Send To"),
                                                      p("Choose from the available options to specify who will receive the emails. Multiple columns may be selected.")),
                                              tags$li(strong("CC"),
                                                      p("Specify an Email address to CC or use a flag for an email CC, if multiple emails are to be specified, they must be separated via \";\"")),
                                              tags$li(strong("Subject Line"),
                                                      p("Write the Subject for the Emails. Flags may be used here.")),
                                              tags$li(strong("Include Documents Made"),
                                                      p("If toggled on, the application will attempt to attach each document listed in the \"DocumentPaths\" column.
                                       If the file does not exist, there will be an error log. If you want to send only additional documents and not merged documents.
                                       Make the documents and toggle the \"Include documents Made\" toggle off.")),
                                              tags$li(strong("Additional Attachments"),
                                                      p("If you wish to send the same attachment(s) to all email recipients, click this button and select the directory containing
                                       the files. Be careful that only the documents that you wish to attach are in this folder as the application
                                       will attempt to all files located there. All documents in that folder will be sent, please create a new file and include all
                                       additional attachments here. This application will only attach these files if the \"Include Additional 
                                       Attachments\" button is toggled on.")),
                                              tags$li(strong("Send on Behalf"),
                                                      p("Toggle on if you wish to send on behalf of someone else. If toggled on the user must specify an email in the adjacent Email field.
                                       The user must give you permission in their Outlook. They must go to \"Account Settings\", the \"Delegate access\", then \"Add\" to
                                       select the user that will be sending the CDAs. All permissions can be given \"None\", except for \"Inbox\". \"Inbox\" should
                                       be set to \"Editor (Can Read, Create, and Modify Items)\".")),
                                              tags$li(strong("Figure"),
                                                      p("A bar graph will appear if it detects any NA's in the Flags used on the Email section. This will show you the frequency of the NA's per Flag. NA's here will not prevent the user from sending an email, however,
		all value substitutions will be made with an empty string \"\". You may toggle this figure on and off by selecting \"Render Plot?\"."))
                                            )
                                            )
                                    )
                                  ),
                              div(class="tab-pane",
                                  'data-value'="Data Table",
                                  id="tab-5000-5",
                                  h3(strong("Data Table Features")),
                                  tags$ul(
                                    tags$li(strong("Column Visibility"), p("Toggle which columns you want to view at a time")),
                                    tags$li(strong("Download"), p("Specify which file format you wish to download from the table. This will only Download what is currently showing on the table.")),
                                    tags$li(strong("Show X entries"), p("Toggle how many entries are visible")),
                                    tags$li(strong("Filtering Fields"),
                                            p("The user may filter the Data Frame by typing into the fields below each Column Header")),
                                    tags$li(strong("Editability"),
                                            p("each value of the Data Table is editable. Double click on an entry to start an edit and click away to finish and update.
	the act of editing the data table saves a new backup.")
                                            )
                                    )
                                  ),
                              div(class="tab-pane",
                                   'data-value'="Backups",
                                   id="tab-5000-6",
                                   h3(strong("Fail safes ---- Preventative Measures")),
                                   p("Any text contained in a flag that does not exactly match the column headers of the imported data frames will
usually be made known to the user. This will usually display as text below the field the flag was used or as 
a graph. At most 12 of the latest versions of the data table used will be stored in rootR directory. Additionally, a few
fields will also backup your last entry if you exit the application. There is currently no method to back up the
template.docx file or your email text entry. The system will not allow you to send emails if there are fields that do not 
                   correspond to a column name.",
                                     p("File backups are stored in Documents\\EZDocs\\backups directory. If the application is taking too long to load,
                     then you may want to delete the backups folder. A new directory will be made when you next lauch the application.",strong("NOTE:"),
                                       " By doing this you will also reset all input defaults."))
                                   ),
                              div(class="tab-pane",
                                  'data-value'="Backups",
                                  id="tab-5000-7",
                                  h3(strong("Errors ---- How to Fix")),
                                  p("There are two types of errors that could occur. The first is systematic errors. These errors 
                   are predictable and checked for in the application. If a systematic error occurs its output
                   will be placed in the current Output directory. The second is random error. These errors 
                   are more difficult to catch simply because we do not yet know what triggering them. If the
                   application fails unexpectedly, then an error log may be viewed in the rootR directory under the error
                   log \"[Documents\\EZDocs\\logs]\"."),
                                  tags$ul(strong("Systematic Errors"),
                                          tags$li(strong("Document Creation:"),
                                                  "Errors are checked in the following order",
                                                  p("File Name > document Flag Replacement > Flag Value"),
                                                  tags$ul(
                                                    tags$li(strong("File Name"), 
                                                            p("This indicates that at least one of the flags 
      used to build the file name was NA for the index (cell on the data frame).
      Rows with this error do not produce any documents, because 
      the program does not know how to name the file. To fix ensure
	there are no NAs in used Columns.")),
                                                    tags$li(strong("Document Flag Replacement"), 
                                                            p("Flag was not properly replaced in program.
      This can often be solved by CTRL+A, CTRL+X, CTRL+V on the template document
      If you receive this error - it is likely it is persistent in", strong("all"), "of the documents made. In the 
      latest version of this application this should not be an issue.")),
                                                    tags$li(strong("Flag Value"),
                                                            p("Indicates a flag was NA in the document. Documents are still 
      made however an empty string is substituted into the document in place of
      the flag. This is not ideal because it may be an important value to your document, please check documents."))
                                                  )),
                                          tags$li(strong("Emails:"), "There is currently only one Error checked when emails are being sent",
                                                  tags$ul(
                                                    tags$li(strong("Document DNE"),
                                                            p("Indicates that the specified file Path stored in the data frame no longer exists or
                                  the application cannot connect to it. In these cases, it may be best to remake your documents as editing file paths by hand is prone to error."))
                                                  )
                                          )
                                  ),
                                  p("If the application becomes unresponsive and/or will not open again after launching the app, it is likely
                   that the program did not close properly. You may confirm if the application ended properly by reading the 
                   log files located in the rootR directory under log. Open the \"error.log\" file and confirm that the last line
                   reads \"Stopping Application\". If not, then open task manager and end any R front-end task or R related process 
                   running. If you do not see any R process running you may end the current processes by opening cmd app via the start
                   menu. Run:"),p( strong("tasklist | findstr Rscript.exe"), style = "color:grey"), p("to verify R is running. If so terminate this process by running:"),
                                  p(strong("taskkill /IM \"Rscript.exe\" /F"), style = "color:grey"),
                                  h3(strong("Understanding Log files")),
                                  p("The log files will be in the rootR directory in the folder named log. These files are
                   helpful with debugging the software as well as viewing random error reports. Almost all actions
                   are recorded in these log files in order to provide some type of trace to the programs state if 
                   it crashes.")
                                  )
                             )
                          )
                 ), #----- end of custom navlistpanel content...
              
                 
                 
               )
      )
    ),
    fluidRow(
      column(12,
             DTOutput("TableOut")
      )
    )
  )
)









server <- function(input, output, session) {
  
  MaxBackup <- 12
  Date_format <- "%m/%d/%Y" #For more info on date codes in R visit: https://www.stat.berkeley.edu/~s133/dates.html
  source("functions.R")
  source("SetupDirBrowsers.R")
  if(!dir.exists(here("backups"))){
    #Create backup folder if it doesnt exist
    dir.create(here("backups","reactiveImage"), recursive = T)
  }
  
  if(!interactive()){
    session$onSessionEnded(function() {
      isolate(saveImage())
      if(dir.exists(here("log"))){
        errlog <- create_latestversion(path = here("log"), pattern = paste0(Sys.Date(),"_error"), device = ".log")
        file.copy(from = here("log/error.log"), to = errlog)
      }
      #isolate(saveMCEcontent(input$EmailEditor))
      print("Stopping Application")
      stopApp()
      # q("no")
    })
  }
  
  #---- setup
  
  #add hidden files that will be assigned in rv :)
  
  
  #print(getwd())
  
  
  #print(str(.Masterdf))
  
  loadImage()
  #load(here("backups/reactiveImage/data.RData"))
  rv <- reactiveValues(Masterdf = .Masterdf,
                       renderTable = FALSE,
                       ColIDs = .ColIDs,
                       detectedIDs = NULL,
                       OutPutFileMessage = "Set File Output With [Flags]",
                       OutPutEmailMessage = "",
                       DocPath = .DocPath,
                       FileNameOutput = .FileNameOutput,
                       SendTo = .SendTo,
                       CCoption = .CCoption,
                       Subject = .Subject,
                       SendOnBehalf = .SendOnBehalf,
                       BehalfEmail = .BehalfEmail,
                       addAttachPath = .addAttachPath,
                       CanRun = .CanRun)
  
  # BehalfEmail
  saveImage <- reactive({
    if(!dir.exists(here("backups/reactiveImage/"))){
      dir.create(here("backups/reactiveImage/"))
    }
    .ColIDs <- isolate(rv$ColIDs)
    .DocPath <- rv$DocPath
    .FileNameOutput <- rv$FileNameOutput
    .SendTo <- rv$SendTo
    .CCoption <- rv$CCoption
    .Subject <- rv$Subject
    .SendOnBehalf <- rv$SendOnBehalf
    .BehalfEmail <- rv$BehalfEmail
    .addAttachPath <- rv$addAttachPath
    .CanRun <- rv$CanRun
    save(.ColIDs,
         .DocPath,
         .FileNameOutput,
         .SendTo,
         .CCoption,
         .Subject,
         .SendOnBehalf,
         .BehalfEmail,
         .addAttachPath,
         .CanRun,
         file = here("backups/reactiveImage/data.RData"))
    
  })
  
  # observeEvent(input$browser,{
  #   browser()
  #   
  #   return(NULL)
  # })
  
  saveBackup <- reactive({
    path <- here("./backups")
    files <- paste0(path, "/",list.files(path =path, pattern = ".*\\.rds"))
    if(length(files)>=MaxBackup){
      oldestFile <- files[which.min(file.mtime(files))]
      file.remove(oldestFile)
    }
    .backup <- rv$Masterdf
    tname <- gsub(pattern = "\\.xlsx$||\\.xls$||\\.csv", replacement = "", input$TableIn$name)
    tim <- gsub(" |:", replacement = "-", Sys.time())
    saveRDS(.backup, file = paste0("./backups/DF_",tname,"_", tim ,".rds"))
  })
  
  # output$OptionsPanel <- renderUI({
  #   numericInput("MaxBackup", "Backup Number", value = MaxBackup)
  # })
  
  CheckDF <- reactive({
    if(!is.null(rv$Masterdf)){
      if(!"DocumentPaths" %in% colnames(rv$Masterdf)){
        rv$Masterdf$DocumentPaths <- character(nrow(rv$Masterdf))
      }
      
      if(!"DocumentsMade" %in% colnames(rv$Masterdf)){
        rv$Masterdf$DocumentsMade <- character(nrow(rv$Masterdf))
      }
      test <- isolate({colnames(rv$Masterdf)%in%rv$ColIDs})
      if(!all(test)){
        rv$ColIDs <- colnames(rv$Masterdf)
      }
      
    }
  })
  
  
  
  shinyDirChoose(input, 'dir', roots = dirList, filetypes = c('', 'txt','docx','csv','pdf','xlsx','xls','tbl','msg'))
  shinyDirChoose(input, 'additionalAttachments', roots = dirList, filetypes = c('', 'txt','docx','csv','pdf','xlsx','xls','tbl','msg'))
  
  # observeEvent(input$debug,{
  #   browser()
  #   
  #   return(NULL)
  # })
  #     
  observe({
    req(input$dir)
    if(is.null(input$dir)||!"path"%in%names(input$dir)){
      return(NULL)
    }
    print("Output Directory selected")
    relpath <-  paste0(dirList[[input$dir$root]],
                       "/",
                       paste(unlist(input$dir$path[-1]),
                             collapse = "/"))
    
    rv$DocPath <- normalizePath(c("C://Users",
                                  paste0(
                                    dirname(relpath), "/",
                                    basename(relpath)
                                  )),
                                winslash = "\\")[2]
    
    print(paste(rv$DocPath))
  })
  
  output$dirDisplay <- renderText({
    paste0("Files to be saved in:\n", p(strong(rv$DocPath), style = "color:blue; word-wrap: break-word; word-break: break-all;"))
  })
  
  output$TableOut <- renderDT({
    #     req(input$TableIn)
    #     req(input$Sheet2use)
    table <- input$TableIn
    sheetOpt <- input$Sheet2use
    #req(rv$Masterdf)
    if(rv$renderTable){
      tabletype <- str_extract(table$name,"\\.csv$|\\.xls|\\.xlsx")
      if(tabletype %in% ".csv"){
        data <- read.csv(file = table$datapath, header = T, stringsAsFactors = F)
      } else if(tabletype %in% c(".xlsx",".xls")){
        data <- as.data.frame(read_excel(path = table$datapath, sheet = sheetOpt))
      } else {
        print("Input formats must be .xls, .xlsx, .csv")
      }
      date_col <- unlist(lapply(lapply(data, class),function(x){any(x %in% c("Date", "POSIXct","POSIXt"))}))
      if(sum(date_col)>0){
        date_col <- names(date_col)[date_col]
        for(i in date_col){
          data[[i]] <- format(data[[i]], Date_format)
        }
      }
      rv$Masterdf <- data
      saveBackup()
      rv$renderTable <- FALSE
    }
    
    test <- isolate({colnames(rv$Masterdf)%in%rv$ColIDs})
    if(!all(test)){
      rv$ColIDs <- colnames(rv$Masterdf)
    }
    CheckDF()
    rv$Masterdf
  }, editable = 'cell', class = c("compact stripe cell-border nowrap hover"), filter = 'top',
  extensions = list('Buttons' = NULL,
                    'FixedColumns' = NULL),
  options = list(scrollX = TRUE,
                 dom = 'Bfrltip',
                 lengthMenu = list(c(10, 25, 50, -1), c('10', '25', '50', 'All')),
                 fixedColumns = TRUE,
                 buttons = list(I('colvis'),
                                 list(extend = c('collection'),
                                      buttons = list(list(extend = 'csv',
                                                          filename = paste0(Sys.Date(),"_Submitted_Export")),
                                                     list(extend = 'excel',
                                                          filename = paste0(Sys.Date(),"_Submitted_Export")),
                                                     list(extend = 'pdf',
                                                          filename = paste0(Sys.Date(),"_Submitted_Export"))),
                                      text = 'Download')
                 )
  )
  )
  observeEvent(input$TableOut_cell_edit, {
    rv$Masterdf <<- editData(data = rv$Masterdf, info = input$TableOut_cell_edit, 'TableOut')
    rv$Masterdf <- FixDTCoerce(rv$Masterdf)
    saveBackup()
  })
  
  
  observeEvent(input$backupLoad, {
    showModal(modalDialog(
      selectInput(inputId = "backupSelect", "Available Backups", choices = list.files(path = "./backups/", pattern = ".*\\.rds"), selected = NULL),
      footer = tagList(
        modalButton("Cancel"),
        actionButton("ok", "OK")
      )
    ))
  })
  
  output$DynamicInput <- renderUI({
    textInput(inputId = "FileNameOutput",
              "Customize New File Name", value = isolate(rv$FileNameOutput))
  })
  
  observe({
    rv$FileNameOutput <- input$FileNameOutput
    print(paste0("Customize New File Name changed to: ", rv$FileNameOutput))
  })
  
  
  observe( {
    #req(input$ChooseTemplate)
    #req(input$Masterdf)
    #browser()
    req(rv$ColIDs, rv$Masterdf, rv$FileNameOutput)
    df <- isolate(rv$Masterdf)
    additionalFlags <- gsub(pattern = "\\[|\\]", replacement = "", x = unlist(str_extract_all(rv$FileNameOutput, pattern = "\\[[^\\]]+\\]")))
    CanRun <- additionalFlags %in% colnames(df)
    if(length(CanRun)!=0&&all(CanRun)){
      test <- na.exclude(df[,additionalFlags])
      if(length(additionalFlags)==1){
        nn <- length(test)
        if(nn!=length(unique(test))){
          rv$OutPutFileMessage <- "Flag Combination Is not Unique. Please add another."
          rv$isUnique <- FALSE
        } else {
          rv$OutPutFileMessage <- "There are no duplicates. All Flags are unique and matched with a column labels on your excel table."
          rv$isUnique <- TRUE
        }
      } else {
        nn <- nrow(test)
        if(nn!=nrow(unique(test))){
          rv$OutPutFileMessage <- "Flag Combination Is not Unique. Please add another."
          rv$isUnique <- FALSE
        } else {
          rv$OutPutFileMessage <- "There are no duplicates. All Flags are unique and matched with a column labels on your excel table."
          rv$isUnique <- TRUE
        }
      }
    } else {
      .num <- additionalFlags[!additionalFlags %in% colnames(df)]
      m <- str_flatten(paste0("\"",.num,"\""), ", ")
      if(length(.num)>1){
        rv$OutPutFileMessage <- paste0("The following flags are not found as column labels in your excel table: ", m)
      } else {
        rv$OutPutFileMessage <- paste0("The following flag is not found as a column label in your excel table: ", m)
      }
      rv$isUnique <- FALSE
    }
    
  })
  
  observe({
    print(paste0("Template Document uploaded: ", input$ChooseTemplate$name))
  })
  
  observe({
    print(paste0("Data Table uploaded: ", input$TableIn$name))
  })
  
  observe({
    print(paste0("Sheet Selected: ", input$Sheet2use))
  })
  
  observe({
    print(paste0("Data Table backup selected: ", input$backupSelect))
  })
  
  #Find any flags not in masterDF - Report them here
  observe({
    # browser()
    df <- isolate(rv$Masterdf)
    emailedit <- input$EmailEditor
    emailedit <- str_remove_all(emailedit, pattern = "\\<[^\\>]+\\>")
    longstr <- c(emailedit,rv$Subject, rv$BehalfEmail, rv$CCoption)
    additionalFlags <- gsub(pattern = "\\[|\\]", replacement = "", x = unlist(str_extract_all(longstr, pattern = "\\[[^\\]]+\\]")))
    additionalFlags <- c(additionalFlags,  rv$SendTo)
    additionalFlags <- unique(additionalFlags)
    CanRun <- additionalFlags %in% colnames(df)
    if(length(CanRun)!=0&&all(CanRun)){
      rv$OutPutEmailMessage <- "All Flags Present in dataframe!"
      rv$CanRun <- TRUE
    } else {
      m <- str_flatten(paste0("\"",additionalFlags[!additionalFlags %in% colnames(df)],"\""), ", ")
      rv$OutPutEmailMessage <- paste0("Not Found in column headers : ", m)
      rv$CanRun <- FALSE
    }
  })
  
  
  output$MessageFileNameOutputDynam <- renderText({
    txt <-rv$OutPutFileMessage
    if(is.null(rv$isUnique)||!rv$isUnique){
      txt <- strong(p(txt,style ="color:red"))
    } else {
      txt <- strong(p(txt), style = "color:green;")
    }
    as.character(txt)
    #textOutput("MessageFileNameOutput")
    
  })
  
  
  output$EmailMessage <- renderText({
    txt <- rv$OutPutEmailMessage
    if(!txt %in%c("All Flags Present in dataframe!")){
      txt <- p(strong(txt), style = "color:red")
    } else {
      txt <- strong(p(txt), style = "color:green;")
    }
    as.character(txt)
    })
  
  observeEvent(input$ok, {
    load(paste0(path,"/",input$backupSelect))
    rv$Masterdf <- .backup
    rv$ColIDs <- colnames(rv$Masterdf)
    CheckDF()
    removeModal()
  })
  
  #   observeEvent(input$ok2, {
  #     removeModal()
  #     updateTextInput(session, inputId = "FileNameOutput",label = "Customize New File Name", value = input$NotUnique)
  #   })
  
  
  output$DetectedIDs <- renderText({
    req(input$ChooseTemplate)
    numIds <- length(rv$ColIDs)
    doctext <- RDCOMExtractText(docpath = input$ChooseTemplate$datapath) #Saves text
    detected <- detectFlags(LongSTR = doctext) 
    InColumnSpace <- detected %in% rv$ColIDs
    if(length(detected)==0){
      str <- p(strong(paste0("No detected flags found in Template docx. Flags should be between \"[\" and \"]\" in plane text. Please Check your Template docx and upload again.")), style="color:red")
    } else{
      if(sum(!InColumnSpace)==0){
        str <- p(paste0(length(detected),
                      " detected flags found in Template docx. All ",
                      length(detected)," detected flags are in the column headers."),style="color:green")
      } else {
        str <- p(strong(paste0(length(detected),
                      " detected flags found in Template docx. ",
                      length(detected)-sum(!(InColumnSpace)),
                      " of ",
                      length(detected),
                      " detected flags are In the column headers. All ",
                      length(detected)," detected flags must be in the column headers. Unused flags: ", str_flatten(paste0("\"",detected[!InColumnSpace],"\""), collapse = ", "))), style="color:red")
      }
      
    }
    rv$detectedIDs <- detected
    as.character(str)
  })
  
  observeEvent(input$Sheet2use,
               {
                 rv$renderTable <- TRUE
               })
  output$SheetSelection <- renderUI({ #Look into this?? ----
    req(input$TableIn)
    #req(input$Sheet2use)
    xls.options <- c("xlsx","xls")
    
    if(grep("\\.xlsx$||\\.xls$", x =input$TableIn$name)==1&&
       format_from_ext(input$TableIn$datapath)%in%xls.options) {
      rv$renderTable <- TRUE
      sheets <- excel_sheets(input$TableIn$datapath)
      selectInput("Sheet2use", label = "Choose Excel Sheet to Load", choices = sheets,selected = sheets[1])
    } else if(grep("\\.csv$", x =input$TableIn$name)==1){
      rv$renderTable <- TRUE
      return(NULL)
    } else {
      rv$renderTable <- FALSE
      textOutput(outputId = "MessageOutput")
    }
  })
  
  
  output$MessageOutput <- renderText({
    "Please Upload an Excel File."
  })
  observeEvent(input$MakeDocuments,{
    if(is.null(rv$DocPath)||!dir.exists(rv$DocPath)){
      if(is.null(rv$DocPath)){
        m <- "Output Directory"
      } else {
        m <- rv$DocPath
      }
      showModal(
        modalDialog(
          p(paste0(m, " has been moved or deleted. Please select a new Output Directory and try again."), style = "color: red;"),
          easyClose = T
        )
      )
      return(NULL)
    } else if(!rv$isUnique){
      showModal(
        modalDialog(
          p(paste0("Flag combination will create duplicates document names. Add an additional Flag to make unique documents."), style = "color: red;"), #style??
          easyClose = T
        )
      )
      return(NULL)
    } else if(is.null(input$ChooseTemplate)){
      showModal(
        modalDialog(
          p(paste0("A template .docx file has not yet been selected."), style = "color: red;"),
          easyClose = T
        )
      )
      return(NULL)
    } else if(is.null(rv$Masterdf)){
      showModal(
        modalDialog(
          paste0("No data frame has been uploaded. Please upload an excel or csv file."),
          easyClose = T
        )
      )
      return(NULL)
    } else if(!all(rv$detectedIDs%in%colnames(rv$Masterdf))){
      showModal(
        modalDialog(
          paste0("Not all detected flags are present in the header of the dataframe. Please either upload a new template docx or data frame."),
          easyClose = T
        )
      )
      return(NULL)
    }
    showModal(
      modalDialog(
        title = "Review",
        fluidRow(
          tags$ul(
            tags$li("Template File: ", input$ChooseTemplate$name),
            tags$li("Output Directory: ", rv$DocPath),
            tags$li("File Name Output: ", rv$FileNameOutput),
            tags$li("Number of Rows in Data Frame: ", nrow(rv$Masterdf))
            )
          ),
        inputPanel(
                 radioButtons(inputId = "Docs2make",
                              label = "Documents To Make",
                              choices = c("both","docx","pdf"),
                              selected = "both", inline = T),
                 fluidRow(
                   checkboxInput("usepass", "Apply Password For Track Changes?", value = F),
                   textInput("docpass", "Password (.docx)")
                 )
                 
        ),
        footer = tagList(
          actionButton("MakeDocuments2", "Start"),
          modalButton("Cancel")
        ),
        easyClose = F
      )
    )
  })
  
  # observe({
  #   input$Docs2make <- input$Docs2make
  #   input$usepass <- input$usepass
  #   input$docpass <- input$docpass
  # })
  
  observeEvent(input$MakeDocuments2,{
    removeModal()
    #browser()
    if(input$usepass&&str_length(input$docpass)!=0){
      pw <- input$docpass
    } else {
      pw <- NULL
    }
    CheckDF()
    df <- rv$Masterdf
    additionalFlags <- gsub(pattern = "\\[|\\]", replacement = "", x = unlist(str_extract_all(rv$FileNameOutput, pattern = "\\[[^\\]]+\\]")))
    usingFlags <- unique(c(rv$detectedIDs, additionalFlags, "DocumentPaths", "DocumentsMade"))
    n <- nrow(df)
    wordApp <- COMCreate("Word.Application") #creates COM object
    er <- character()
    indx <- numeric()
    whichflag <- character()
    j <- 1
    withProgress(message = "Making Documents", value = 0/n, expr = {
      for(i in 1:nrow(df)){
        print(paste0("Attempting Documents Row: ",i))
        tmpdata <- df[i,usingFlags]
        test <- tmpdata[1, additionalFlags]
        if(length(additionalFlags)==1){
          test <- is.na(test)
        } else{
          test <- apply(test, 2, FUN = is.na) #Test if we can even write the file
        }
        if(any(test)){
          incProgress(1/n, detail = paste("Row", i, "contains NAs in Output File flags ... skipping"))
          er[j] <- "File Name"
          indx[j] <- i
          whichflag[j] <- paste0(additionalFlags[test],collapse = ", ")
          j <- j+1
          print(paste("Skipping Row ", i, ", File Name Error"))
        } else { # if(str_length(tmpdata$DocumentsMade)==0||length(grep(tmpdata$DocumentsMade, input$ChooseTemplate$name, fixed = TRUE))==0)   ....equals not found ... Add overwrite toggle
          incProgress(1/n, detail = paste("Creating Document for row #", i))
          print(paste0("Making Documents Row: ",i))
          docpaths <- RDCOMFindReplace(flags = usingFlags,
                                       data = tmpdata,
                                       docpath = input$ChooseTemplate$datapath,
                                       targetdir = paste0(rv$DocPath,
                                                          "/",
                                                          rv$FileNameOutput),
                                       wordApp = wordApp,
                                       makeWhich = input$Docs2make,
                                       pw = pw)
          if(!is.null(docpaths[["error"]])){
            er[j] <- docpaths[["error"]]
            indx[j] <- i
            j <- j+1
          }
          docpaths <- docpaths[["str"]]
          if(str_length(df[i,]$DocumentPaths)==0||is.na(df[i,]$DocumentPaths)){ #if DocuPath isnt filled, record path
            df[i,"DocumentPaths"] <- docpaths
          } else if(length(grep(x = df[i,]$DocumentPaths, pattern = docpaths, fixed = TRUE))==0){ #if docpaths isnt already there, concatonate path in. (otherwise do nothing because file was overwritten)
            df[i,"DocumentPaths"] <- str_flatten(c(df[i,]$DocumentPaths, docpaths), "|.|")
          } else {
            print("Documents have been overwritten!")
          }
          if(str_length(df[i,]$DocumentsMade)==0||is.na(df[i,]$DocumentsMade)){
            df[i,"DocumentsMade"] <- input$ChooseTemplate$name
          } else if(length(grep(x = df[i,]$DocumentsMade, pattern = input$ChooseTemplate$name, fixed = TRUE))==0){
            df[i,"DocumentsMade"] <- str_flatten(c(df[i,]$DocumentsMade, input$ChooseTemplate$name),"|.|")
          }
          
          test <- tmpdata[1, rv$detectedIDs]
          if(length(rv$detectedIDs)==1){
            test <- is.na(test)
          } else{
            test <- apply(test, 2, FUN = is.na)
          }
          if(any(test)){
            s <- sum(test)
            if(s==1){
              er[j] <- "Flag Value"
              whichflag[j] <- paste0(rv$detectedIDs[test],collapse = ", ")
              indx[j] <- i
              j <- j+1
            } else {
              er[j] <- "Flag Value"
              whichflag[j] <- paste0(rv$detectedIDs[test],collapse = ", ")
              indx[j] <- i
              j <- j+1
            }
            print("Flag Value")
          } 
        }
      }
      
      wordApp$Quit() #quit wordApp
      ERROR <- data.frame(ErrorType = er, RowNumber = indx, flags = whichflag)
      rv$Masterdf[,usingFlags] <- df[,usingFlags]
      saveBackup()
      print("Done making Documents")
    })
    rm(wordApp)
    if(nrow(ERROR)>0){
      rv$DocumentsWarning <- ERROR
      showModal(
        modalDialog(strong("WARNING", style = "color: red;"),
                    p("Please check the documents below for their respective errors!", style = "color: red;"),
                    tags$ul(
                      tags$li(strong("File Name"), 
                              p("This indicates that at least one of the flags 
      used to build the file name was NA for the index (cell on the data frame).
      Rows with this error do not produce any documents.")),
                      tags$li(strong("Flag Value"),
                              p("Indicates a flag was NA in the document. Documents are still 
      made however an empty string is substituted into the document in place of
      the flag. This is not ideal because it may be an important value to your document, please check documents."))
                    ),
                    p("For more information check the error help page."),
                    wellPanel(
                                     dataTableOutput("DocumentsWarning"), style="height:350px;overflow-y:scroll;"
                                     #scrollCollapse:true;
                      ),
                    easyClose = F
                    )
        )
      
      
    }
    

    
    
  })
  
  output$DocumentsWarning <- renderDT({
    rv$DocumentsWarning
  }, class = c("compact stripe cell-border nowrap hover"),
  extensions = list('Buttons' = NULL,
                    'FixedHeader' = NULL),
  options = list(dom = 'Bfrltip',
                 fixedHeader = TRUE,
                 pageLength = 10,
                 buttons = c('copy', 'csv', 'excel', 'pdf'),
                 paging = T), server = F
  )
  
  # output$emailpreview1 <- renderUI({
  #     req(input$EmailEditor)
  #     HTML(enc2utf8(input$EmailEditor))})
  #This will first test that officer package is working and replaces values and saves documents.
 
  observeEvent(input$SendEmails,{
    
    if(any(rv$nadf$NumNA>0)){
      df <- rv$nadf[rv$nadf$NumNA>0,]
      rownames(df) <- 1:nrow(df)
      colnames(df) <- c("Flags","NumberOfNAs")
      rv$emaildftable <- df
      showModal(
        modalDialog(
          title = strong("Warning",style="color:red;"),
          p("Some cells in the dataframe contain NA's/empty values. 
            If you continue they will be replaced with nothing. Do you wish to continue?"),
          wellPanel(
            dataTableOutput("EmailWarning"), style="height:350px;overflow-y:scroll;"
          ),
          footer = tagList(
            actionButton("SendEmails2", "Continue"),
            modalButton("Cancel")
          ),
          easyClose = F
        )
      )
    } else {
      SendTheEmails()
    }
    
  }) 
  
  output$EmailWarning <- renderDT({
    rv$emaildftable
  }, class = c("compact stripe cell-border nowrap hover"),
  extensions = list('Buttons' = NULL,
                    'FixedHeader' = NULL),
  options = list(dom = 'Bfrltip',
                 fixedHeader = TRUE,
                 pageLength = 10,
                 buttons = c('copy', 'csv', 'excel', 'pdf'),
                 paging = T), server = F)
  
  
  observeEvent(input$SendEmails2, {
    removeModal()
    SendTheEmails()
  })
  
  SendTheEmails <- reactive({
    req(input$EmailEditor)
    if(rv$CanRun&&!(is.null(rv$Subject))){
      #req(rv$BehalfEmail)
      df <- rv$Masterdf
      n <- nrow(df)
      #Get flags that will be used for email
      #check input - #EmailEditor, Subject, IncludeAttachments, SendOnBehalf, BehalfEmail
      longstr <- paste(input$EmailEditor,rv$Subject, rv$SendTo, rv$CCoption, rv$BehalfEmail)
      additionalFlags <- gsub(pattern = "\\[|\\]",
                              replacement = "",
                              x = unlist(str_extract_all(longstr,
                                                         pattern = "\\[[^\\]]+\\]")))
      sub <- rv$Subject
      Body <- input$EmailEditor
      To <- rv$SendTo
      CC <- rv$CCoption
      CCisFixed <- str_detect(CC, pattern = "\\[|\\]", negate = TRUE)
      additionalFlags <- unique(c(additionalFlags, "DocumentPaths", To))
      additionalFlags <- additionalFlags[additionalFlags %in% colnames(df)]
      OutApp <- COMCreate("Outlook.Application")
      er <- character()
      indx <- numeric()
      docID <- character()
      k <- 1
      print("Starting to Send Emails")
      print(paste0("Subject: ", sub))
      print(paste0("To: ", To))
      print(paste0("CC: ", CC))
      if(rv$SendOnBehalf){
        print(paste0("Send on Behalf: ",rv$BehalfEmail))
      } else {
        print(paste0("Send on Behalf: ", rv$SendOnBehalf))
      }
      print(paste0("Including Documents in DocumentPaths column: ",input$IncludeAttachments))
      print(paste0("Including Additional Files: ", input$IncludeAddAttachment))
      withProgress(message = "Sending Emails", value = 0/n, expr = {
        for(i in 1:nrow(df)){
          print(paste("Emails: Starting row ",i))
          tmpdata <- df[i, additionalFlags]
          Totmp <- sub_flags(To, data = tmpdata)
          if(CCisFixed){
            CCtmp <- CC
          } else {
            CCtmp <- gsub(pattern = "\\[|\\]",replacement = "", x = str_extract_all(string = CC, pattern = "\\[.+?\\]")[[1]])
            CCtmp <- sub_flag_in_str(x = CC, flags = CCtmp, data = tmpdata)
          }
          
          #CheckFor NA or empty strings
          .t <- str_length(Totmp)
          Totmp <- Totmp[!(.t==0|is.na(.t))]
          if(length(Totmp)>1){
            Totmp <- paste(Totmp, collapse = ";")
          } else if(length(Totmp)==0){
            Totmp <-""
          }
          
          subtmp <- sub_flag_in_str(x = sub, flags = additionalFlags, data = tmpdata)
          Bodytmp <- sub_flag_in_str(x = Body, flags = additionalFlags, data = tmpdata)
          
          outMail = OutApp$CreateItem(0)
          outMail$GetInspector()
          signature <- outMail[["HTMLBody"]]
          outMail[["to"]] = Totmp
          CCtmp <- gsub(pattern = " ", replacement = "", x = CCtmp, fixed = T)
          if(str_length(CCtmp)!=0&&!CCtmp %in% c(",")){
            outMail[["Cc"]] = CCtmp
          }
          outMail[["subject"]] = subtmp
          outMail[["HTMLBody"]] = paste0("<p>",Bodytmp,signature,"</p>")
          print("attempting to attach documents")
          if(rv$SendOnBehalf){
            outMail[["sentonbehalfofname"]] <- rv$BehalfEmail
          }
          if(input$IncludeAttachments){
            docPaths <- tmpdata[["DocumentPaths"]] #sub_flags(x = "DocumentPaths", data = tmpdata)
            docPaths <- unlist(str_split(string = docPaths, pattern = "\\|\\.\\|"))
            if(input$Docs2include %in% "docx"){
              docPaths <- docPaths[grep(pattern = "\\.docx$", x = docPaths)]
            } else if(input$Docs2include %in% "pdf") {
              docPaths <- docPaths[grep(pattern = "\\.pdf$", x = docPaths)]
            }
            for(j in 1:length(docPaths)){
              if(file.exists(docPaths[j])){
                outMail[["attachments"]]$Add(docPaths[j])
              } else if(docPaths[j]!="") {
                er[k] <- "Document DNE"
                docID[k] <- docPaths[j]
                indx[k] <- i
                k <- k+1
              }
            }
            
          }
          if(input$IncludeAddAttachment&&!is.null(rv$addAttachPath)){
            addfiles <- !dir.exists(paths = paste0(rv$addAttachPath,"\\",list.files(rv$addAttachPath)))
            addfiles <- paste0(rv$addAttachPath,"\\",list.files(rv$addAttachPath))[addfiles]
            for(j in 1:length(addfiles)){
              if(file.exists(addfiles[j])){
                outMail[["attachments"]]$Add(addfiles[j])
              } else {
                er[k] <- "Document DNE"
                docID[k] <- addfiles[j]
                indx[k] <- i
                k <- k+1
              }
            }
          }
          if(str_length(outMail[["to"]])==0){
            progmessage <-  paste0(i,": No sending address. Saved as draft!")
            outMail$Save()
          } else {
            progmessage <- paste("Sent Email for row #", i)
            print(paste0("Sending Email for row ", i,"."))
            outMail$Send()
          }
          
          
          incProgress(1/n, detail = progmessage)
        }
        
        if((input$IncludeAttachments||input$IncludeAddAttachment)&&length(er)>0){
          if(is.null(rv$DocPath)||!dir.exists(rv$DocPath)){
            print("Document path is not set. Setting to rootR/logs")
            docpaths <- normalizePath(c("C:/Users/",here("log")), winslash = "\\")[[2]]
          } else {
            docpaths <- rv$DocPath
          }
          erPath <- paste0(docpaths,
                           "\\",
                           Sys.Date(),"_Attachment_ERROR_Log.csv")
          if(file.exists(erPath)){
            erPath <- create_latestversion(path = docpaths,
                                           pattern = paste0(Sys.Date(),"_Attachment_ERROR_Log"),
                                           device = ".csv")
          }
          ERROR <- data.frame(ErrorType = er, Index = indx, docPath=docID)
          write.csv(x = ERROR, file = erPath)
        }
        
      })
      print("Done sending emails!")
      rm(outMail, OutApp)
    } else{
      showModal(
        modalDialog(
          p("Please be sure that all flags are present in the data frame before sending Emails."),
          p(strong(rv$OutPutEmailMessage), style = "color:red"),
          easyClose = T
        )
      )
    }
  })
  
  
  output$NAflags <- renderUI({
    
    if(is.null(rv$Masterdf)){
      rendertext <- T
    } else {
      rendertext <- F
    }
    if(is.null(input$renderplot)){
      plot <- F
    } else {
      plot <- input$renderplot
    }
    if(rendertext||!plot){
      htmlOutput(outputId = "EmailMessage")
    } else {
      plotOutput("NAplot",)
    }
    
  })
  
  observe({
    req(input$EmailEditor, rv$Masterdf)
    longstr <- paste(input$EmailEditor,rv$Subject, rv$CCoption, rv$BehalfEmail)
    additionalFlags <- gsub(pattern = "\\[|\\]",
                            replacement = "",
                            x = unlist(str_extract_all(longstr,
                                                       pattern = "\\[[^\\]]+\\]")))
    if(!is.null(rv$SendTo)){
      additionalFlags <- c(additionalFlags, rv$SendTo)
    }
    additionalFlags <- unique(additionalFlags)
    df <- rv$Masterdf
    .df <- df[,colnames(df) %in% additionalFlags]
    if(sum(colnames(df) %in% additionalFlags)==1){
      nadf <- sum(is.na(.df)|(str_length(.df)==0))
    } else {
      nadf <- apply(.df, 2, function(x){sum(is.na(x)|(str_length(x)==0))})
    }
    NumNonz <- nadf!=0
    nadf <- data.frame(Flags = additionalFlags[additionalFlags %in% colnames(df)], NumNA=nadf)
    if(sum(!additionalFlags %in% colnames(df))>0){
      addfakes <- data.frame(Flags = additionalFlags[!additionalFlags %in% colnames(df)], NumNA =0)
      nadf <- rbind(nadf, addfakes)
    }
    rv$nadf <- nadf
  })
  
  output$NAplot <- renderPlot({
    m <- rv$OutPutEmailMessage
    NumNonz <- rv$nadf$NumNA!=0
    ggplot(rv$nadf, aes(x=as.factor(Flags), y = as.numeric(NumNA))) +
      geom_bar(stat = "identity", fill = "blue") +
      labs(x = "\nFlags",
           y = "Number of NA\n",
           title = m, subtitle = paste0(sum(NumNonz),
                                        " out of ",
                                        length(NumNonz),
                                        " columns contain NA's")) +                                                       
      theme_classic() 
    
    
  })
  

  observe({
    req(input$additionalAttachments)
    if(is.null(input$additionalAttachments)||!"path"%in%names(input$additionalAttachments)){
      return(NULL)
    }
    #browser()
    print("Additional Documents Path Selected")
    relpath <-  paste0(dirList[[input$additionalAttachments$root]],
                       "/",
                       paste(unlist(input$additionalAttachments$path[-1]),
                             collapse = "/"))
    
    rv$addAttachPath <- normalizePath(c("C://Users",
                                  paste0(
                                    dirname(relpath), "/",
                                    basename(relpath)
                                  )),
                                winslash = "\\")[2]
    print(paste(rv$addAttachPath))
  })
  
  output$Attachments_info <- renderText({
    if(input$IncludeAddAttachment&&!is.null(rv$addAttachPath)){
      n <- dir.exists(paths = paste0(rv$addAttachPath,"\\",list.files(rv$addAttachPath)))
      n <- sum(!n)
      print(paste0(n," files to be attached to all emails. Files located in:\n", rv$addAttachPath))
      paste0(n," files to be attached to all emails. Files located in:\n", p(strong(rv$addAttachPath), style = "color:blue;word-wrap: break-word; word-break: break-all;"))
    } else {
      return(NULL)
    }
  })
  
  observe({
    print(paste0("Include Additional Attachments toggle set to: ", input$IncludeAddAttachment))
  })
  
  observe({
    print(paste0("Include Documents Made toggle set to: ", input$IncludeAttachments))
  })
  
  output$EmailPanel <- renderUI({
    # req(rv$ColIDs,
    #     rv$SendTo,
    #     rv$CCoption,
    #     rv$Subject,
    #     rv$SendOnBehalf,
    #     rv$BehalfEmail)
    column(6,
           fluidRow(
             column(6,
                    selectInput("SendTo", "Send To:", choices = rv$ColIDs, selected = isolate(rv$SendTo), multiple = TRUE)
             ),
             column(6,
                    textInput(inputId = "CCoption",
                              label = "CC", value = isolate(rv$CCoption))
             )
           ),
           textInput("Subject",
                     "Subject Line",
                     value = isolate(rv$Subject)#"Input Your Subject"
           ),
           fluidRow(
             column(6,
                    checkboxInput("IncludeAttachments",
                                  "Include Documents Made",
                                  value = TRUE),
                    radioButtons(inputId = "Docs2include",
                                 label = "Documents To Include",
                                 choices = c("both","docx","pdf"),
                                 selected = "both")
                    ),
             column(6, align = "center",
                    checkboxInput("IncludeAddAttachment",
                                  "Include Additional Attachments?",
                                  value = FALSE),
                    shinyDirButton(id = "additionalAttachments",
                                   "Additional Attachments",
                                   title = "Attachments For All"),
                    htmlOutput(outputId = "Attachments_info")
                    )
           ),
           
           
           fluidRow(
             column(6,
                    checkboxInput("SendOnBehalf",
                                  "Send On Behalf?",
                                  value = isolate(rv$SendOnBehalf)),
                    checkboxInput(inputId = "renderplot", label = "Render Plot?")
             ),
             column(6,
                    textInput("BehalfEmail",
                              "Email", value = isolate(rv$BehalfEmail))
             )
           ),
           
    )
  })
  
  output$EmailOutputsignature <- renderText({
    OutApp <- COMCreate("Outlook.Application")
    outMail = OutApp$CreateItem(0)
    outMail$GetInspector()
    signature <- outMail[["HTMLBody"]]
    outMail$Close(1)
    rm(outMail, OutApp)
    signature
  })
  observe({ 
    if(!is.null(input$SendTo)){
      if(is.null(rv$SendTo)||
         length(rv$SendTo)!=length(input$SendTo)||
         !all(input$SendTo %in% rv$SendTo)){
        rv$SendTo <- input$SendTo
        print(paste0("Send To selected: ",paste0(rv$SendTo, collapse = ", ")))
      }
    }
  })
  observe({
    if(!is.null(input$CCoption)){
      if(is.null(rv$CCoption)||rv$CCoption!=input$CCoption){
        rv$CCoption <- input$CCoption
        print(paste0("CC field updated to: ", rv$CCoption))
      }
    }
  })
  observe({
    if(!is.null(input$Subject)){
      if(rv$Subject!=input$Subject){
        rv$Subject <- input$Subject
        print(paste0("Subject field updated to: ", rv$Subject))
      }
    }
    
  })
  observe({
    if(!is.null(input$SendOnBehalf)){
      rv$SendOnBehalf <- input$SendOnBehalf
      print(paste0("Send on behalf toggle set to: ", rv$SendOnBehalf))
    }
  })
  observe({
    if(!is.null(input$BehalfEmail)){
      rv$BehalfEmail <- input$BehalfEmail
      print(paste0("Send on Behalf Email field updated to: ", rv$BehalfEmail))
    }
  })
  
  
  observeEvent(input$ExportAllData,{
    if(!is.null(rv$DocPath)){
      showModal(
        modalDialog(
          title = "Export Data",
          paste0("Output Directory:\n",rv$DocPath),
          inputPanel(
            textInput(inputId = "DataOutName",
                      label = "Data Document Name",
                      value = paste0("DF_",Sys.Date()))
          ),
          footer = tagList(
            actionButton("saveExcel", "Save"),
            modalButton("Cancel")
          ),
          easyClose = F
        )
      )
    }
    
  })
  
  observeEvent(input$saveExcel, {
    #Could do some checks here
    if(!is.null(rv$DocPath)&&dir.exists(rv$DocPath)){
      write.csv(rv$Masterdf, file = paste0(rv$DocPath,"\\",input$DataOutName,".csv"), row.names = F)
      removeModal()
    } else {
      removeModal()
      showModal(
        modalDialog(
          paste0(rv$DocPath, " has been moved or deleted. Please select a new Output Directory and try again."),
          easyClose = T
        )
      )
    }
  })
  
  
  
  
}





shinyApp(ui, server)

#Sys.setenv("Tar"="internal")
# RInno::create_app(app_name = "EZDocs",
#                   app_dir = here::here(),
#                   dir_out = "Installer", pkgs = c("shiny","shinyFiles",
#                                                   "DT","readxl","stringr",
#                                                   "RJSONIO","ggplot2","here"),
#                   remotes = c("omegahat/RDCOMClient","mul118/shinyMCE"),
#                   include_R = T,
#                   user_browser = "chrome",
#                   app_icon = "EZDocs.ico",
#                   app_version = "0.1.10"
# 
# )
# RInno::compile_iss()

