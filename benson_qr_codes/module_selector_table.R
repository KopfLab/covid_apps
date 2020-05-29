# Selector table

#' Selector table server
#'
#' This generates an rhandson table for selecting items to include in downstream operations.
#'
#' @param id_column name of the ID column - make a rownumber or concatenated column if there is no unique identificer column otherwise
#' @param row_column which column to use as the "row number", by default same as the id_column
#' @param column_select dplyr select statement to choose displayed columns and headers, if not provided (=NULL), will display data table exactly as is (minus the id column which becomes the row number)
#' @param page_lengths page length options, first one will be selected
#' @param initial_page_length initially selected page length, first entry of the page_lengths by default
#' @param dom the available table control elements and their order
#' @param selector_buttons whether the selector buttons are present
#' @family selector table module functions
selector_table_server <- function(input, output, session, id_column, row_column = id_column, column_select = NULL,
                                page_lengths = list( c(5, 10, 20, -1),  c("5", "10", "20", "All")),
                                initial_page_length = page_lengths[[1]][1], dom = "fltip", selector_buttons = TRUE) {

  # safety checks
  stopifnot(!missing(id_column))

  # namespace
  ns <- session$ns

  # column selection
  column_select_quo = rlang::enquo(column_select) # what should be displayed

  # reactive values
  values <- reactiveValues(
    table = NULL, # what is available
    selected = c(), # what is selected
    update_selected = 0, # trigger selection update (circumventing circular triggers with user selection)
    page_length = initial_page_length, # selected page length
    display_start = 0, # which display page to start on
    search = "", # search term
    order = list() # ordering information
  )

  # render table
  output$selection_table <- DT::renderDataTable({
    values$table # trigger with update of the data table (whole re-render required for client-side)
    isolate({
      validate(need(values$table, "None available."))
      module_message(ns, "info", "TABLE (re-) rendering with ",
                     nrow(values$table), " rows")
      # prepare data
      row_names <- values$table[[row_column]]
      if (!rlang::quo_is_null(column_select_quo)) {
        # take specified selection
        df <- dplyr::select(values$table, !!column_select_quo) %>% as.data.frame()
      } else {
        # remove id and row columns, rest stays the same
        df <- as.data.frame(values$table)
        df[[id_column]] <- NULL
        df[[row_column]] <- NULL
      }
      rownames(df) <- row_names
      # make sure selection staes the same
      update_selected()
      # generate data table
      DT::datatable(
        data = df,
        options = list(
          ordering = values$order,
          pageLength = values$page_length,
          search = list(regex = FALSE, caseInsensitive = TRUE, search = values$search),
          displayStart = values$display_start,
          lengthMenu = page_lengths,
          searchDelay = 100,
          dom = dom,
          # save state to get ordering and other information
          stateSave = TRUE,
          # disable the automatic state reload to avoid issues between different table instances
          stateLoadParams = DT::JS("function (settings, data) { return false; }")
        )
      )
    })},
    server = FALSE # client side usually faster to use with small selector tables
    # NOTE: should this ever change, it will require updates to the way the data
    # table is replaced too (using the dataTableProxy).
  )

  # trigger selection updates
  update_selected <- function() values$update_selected <- values$update_selected + 1
  observeEvent(values$update_selected, {
    selections <- which(values$table[[id_column]] %in% values$selected)
    module_message(ns, "info", "TABLE selecting ", length(selections), " row(s)")
    proxy <- DT::dataTableProxy("selection_table")
    DT::selectRows(proxy, selections)
  }, ignoreInit = TRUE)

  # save state
  observeEvent(input$selection_table_state, {
    isolate({
      module_message(ns, "info", "TABLE saving state")
      values$page_length <- input$selection_table_state$length
      values$display_start <- input$selection_table_state$start
      values$search <- input$selection_table_state$search$search
      values$order <- input$selection_table_state$order
    })
  }, ignoreInit = TRUE)

  # save selection
  observeEvent(input$selection_table_rows_selected, {
    selections <- input$selection_table_rows_selected
    if (!identical(values$selected, selections) && !is.null(values$table) && nrow(values$table) > 0) {
      values$selected <- if (is.null(selections)) c() else values$table[[id_column]][selections]
    }
  }, ignoreNULL = FALSE, ignoreInit = TRUE)

  # selector buttons
  if (selector_buttons) {
    # select all that match the current filter
    observeEvent(input$select_all, {
      values$selected <- unique(c(values$selected, values$table[[id_column]][input$selection_table_rows_all]))
      update_selected()
    })

    # deselect all
    observeEvent(input$deselect_all, {
      values$selected <- c()
      update_selected()
    })

    # button visibility
    observe({
      toggle("select_all", condition = !is.null(values$table))
      toggle("deselect_all", condition = !is.null(values$table))
    })
  }

  # functions
  # WARNING: only use inital_selection when first setting the table, not when re-setting it!
  set_table <- function(table, initial_selection = c()) {
    isolate({
      if (is.null(table) || is.null(values$table) || !identical(table, values$table)) {
        initial <- is.null(values$table)
        values$table <- table
        if (initial && length(initial_selection) > 0)
          set_selected(initial_selection)
      }
    })
  }

  set_selected <- function(selected) {
    isolate({
      if (!identical(selected, values$selected) && (length(selected) > 0 || length(values$selected) > 0)) {
        values$selected <- selected
        update_selected()
      }
    })
  }

  get_selected <- eventReactive(values$selected, {
    # make sure all returned selected are valid
    if (length(values$selected) == 0) return(c())
    else return(values$selected[values$selected %in% values$table[[id_column]]])
  }, ignoreNULL = FALSE)

  get_selected_items <- eventReactive(values$selected, {
    # get the actual table items that are selected
    values$table[values$table[[id_column]] %in% values$selected, ]
  }, ignoreNULL = FALSE)

  are_all_selected <- eventReactive(values$selected, {
    setequal(get_selected(), values$table[[id_column]])
  }, ignoreNULL = FALSE)

  get_table_nrow <- eventReactive(values$table, {
    nrow(values$table)
  })

  # return functions
  list(
    set_table = set_table,
    set_selected = set_selected,
    get_selected = get_selected,
    get_selected_items = get_selected_items,
    get_table_nrow = get_table_nrow,
    are_all_selected = are_all_selected
  )
}


#' Selector table UI
#' @family selector table module functions
selector_table_ui <- function(id) {
  ns <- NS(id)
  DT::dataTableOutput(ns("selection_table"))
}

#' Selector table buttons
#' @family selector table module functions
selector_table_buttons_ui <- function(id) {
  ns <- NS(id)
  tagList(
    tooltipInput(actionButton, ns("select_all"), "Select all",
                 icon = icon("check-square-o"),
                 tooltip = "Select all items that match the current search in addition to those already selected."),
    spaces(1),
    tooltipInput(actionButton, ns("deselect_all"), "Deselect",
                 icon = icon("square-o"),
                 tooltip = "Deselct all items (irrespective of the search).")
  )
}
