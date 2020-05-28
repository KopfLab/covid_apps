library(rlang)
library(dplyr)
library(tidyr)
library(purrr)
library(stringr)
library(readxl)
library(officer)

library(shiny)
library(shinyjs)
library(shinyBS)
library(DT)

source("qr_code_funcs.R")
source("shiny_utils.R")
source("shiny_module_selector_table.R")

# app
shinyApp(
  ui = fluidPage(
    shinyjs::useShinyjs(),
    h1("Generate QR Codes for your building"),
    h4("Instructions available", a(href = "https://docs.google.com/document/d/1n42-g-CoqoNPL8JOfx1OAykIDCj9lUEGy_fTGKVIG4E/edit?usp=sharing", "at this link")),
    h2("Pre-filled Form Link"),
    textInput(inputId = "prefilled_form_link", label = NULL, placeholder = "Paste link"),
    h5(
      style = "color: red;", 
      textOutput("form_url"),
      textOutput("room_field_id"),
      textOutput("direction_field_id")
    ),
    
    # rooms
    h2("Select Rooms"),
    downloadButton("rooms_template", "Download Rooms List Template"),
    h4("Upload Custom Rooms List (.xlsx file based on the template)"),
    fileInput("rooms_upload", NULL, multiple = FALSE, accept = c(".xlsx")),
    h4(style = "color: red;", textOutput("rooms_info")),
    selector_table_ui("rooms"),
    inline(
      actionButton("refresh_rooms", "Refresh Rooms", icon = icon("refresh")),
      spaces(1),
      selector_table_buttons_ui("rooms"),
    ),
    
    # qr code generateion
    h2("Select Word Template"),
    downloadButton("word_template", "Download Word Template"),
    h4("Upload Custom Template (.docx file based on the template)"),
    fileInput("template_upload", NULL, multiple = FALSE, accept = c(".docx")),
    h4(style = "color: red;", textOutput("template_info")),  
    downloadButton("generate", "Generate QR Codes for Selected Rooms")
  ),
  server = function(input, output, session) {
    
    # reactive values ====
    values <- reactiveValues(
      form_url = NULL,
      room_field_id = NULL,
      direction_field_id = NULL,
      rooms_path = "template.xlsx",
      template_path = "template.docx"
    )
    
    # form link ====
    observeEvent(input$prefilled_form_link, {
      req(nchar(input$prefilled_form_link) > 0)
      form_values <- parse_prefilled_form_url(input$prefilled_form_link)
      values$form_url <- form_values$form_url
      values$room_field_id <- form_values$room_field_id
      values$direction_field_id <- form_values$direction_field_id
    })
    output$form_url <- renderText({
      if (is.null(values$form_url)) ""
      else if (is.na(values$form_url)) "Form URL could NOT be determined from the provided link"
      else paste("Form URL:", values$form_url)
    })
    output$room_field_id <- renderText({
      if (is.null(values$room_field_id)) ""
      else if (is.na(values$room_field_id)) "Room Field ID could NOT be determined from the provided link"
      else paste("Room Field ID:", values$room_field_id)
    })
    output$direction_field_id <- renderText({
      if (is.null(values$direction_field_id)) ""
      else if (is.na(values$direction_field_id)) "Direction Field ID could NOT be determined from the provided link"
      else paste("Room Field ID:", values$direction_field_id)
    })
    
    # rooms =====
    output$rooms_template <- downloadHandler(
      filename = "template.xlsx",
      content = function(file) {
        file.copy("template.xlsx", file)
      }
    )
    observeEvent(input$rooms_upload, {
      values$rooms_path <- input$rooms_upload$datapath
    })
    get_rooms <- reactive({
      tryCatch(readxl::read_excel(values$rooms_path),
               error = function(e) {
                 paste0("cannot open file: ", e$message)
               })
    })
    output$rooms_info <- renderText({
      if (is.data.frame(get_rooms()) && "Room" %in% names(get_rooms())) {
        ""
      } else if (is.data.frame(get_rooms()) && !"Room" %in% names(get_rooms())) {
        "no 'Room' column in the spreadsheet"
      } else if (is.character(get_rooms())) {
        get_rooms()
      } else {
        "cannot retrieve rooms"
      }
    })
    rooms <- callModule(selector_table_server, "rooms", id_column = "Room", initial_page_length = 10)
    observe({
      if (is.data.frame(get_rooms()) && "Room" %in% names(get_rooms()))
        rooms$set_table(get_rooms())
      else
        rooms$set_table(tibble(Room = character(0)))
    })

    # template =====
    output$word_template <- downloadHandler(
      filename = "template.docx",
      content = function(file) {
        file.copy("template.docx", file)
      }
    )
    observeEvent(input$template_upload, {
      values$template_path <- input$template_upload$datapath
    })
    get_template_path <- reactive({
      tryCatch({
        officer::read_docx(values$template_path)
        values$template_path
      },
      error = function(e) {
        paste0("cannot open file: ", e$message)
      })
    })
    output$template_info <- renderText({
      if (!file.exists(get_template_path()))
        get_template_path()
      else
        ""
    })
    
    # generate button ====
    observe({
      disabled <- is.null(values$form_url) || is.na(values$form_url) ||
        is.null(values$room_field_id) || is.na(values$room_field_id) ||
        is.null(values$direction_field_id) || is.na(values$direction_field_id) ||
        length(rooms$get_selected()) == 0 || !file.exists(get_template_path())
      shinyjs::toggleState("generate", !disabled)
    })
    
    # download QR codes =====
    output$generate <- downloadHandler(
      filename = reactive(paste0(paste(rooms$get_selected(), collapse = "_"), "_qr_codes.docx")),
      content = function(file) {
        withProgress(
          generate_qr_codes_doc(
            template_path = values$template_path,
            rooms = rooms$get_selected(), 
            form_url = values$form_url, 
            room_field_id = values$room_field_id, 
            direction_field_id = values$direction_field_id,
            data = rooms$get_selected_items(),
            save_path = file
          ), 
          min = 0, max = 1, 0.5,
          message = sprintf("Generating QR codes for %d rooms...", isolate(length(rooms$get_selected()))), 
          detail = isolate(paste(rooms$get_selected(), collapse = "_"))
        )
      }
    )
    
  }
)
