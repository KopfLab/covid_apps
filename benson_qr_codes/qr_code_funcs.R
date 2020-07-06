# parse pre-filled form URL
# @param url prefilled form URL with 'room' in the room key and 'enter' or 'exit' in the direction field
# @return list with form_url, room_field_id and direction_field_id
parse_prefilled_form_url <- function(url) {
  reg_exp <- "^(.*)/viewform\\??(.*)$"
  if(!stringr::str_detect(url, reg_exp)) {
    warning("does not look like a prefilled URL, missing 'viewform'", call. = FALSE)
    return (list(
      form_url = NA_character_,
      room_field_id = NA_character_,
      direction_field_id = NA_character_
    ))
  }
  reg_exp_matches <- stringr::str_match(url, reg_exp)
  
  # parse entries
  params <- tibble(
    param = stringr::str_split(reg_exp_matches[,3], fixed("&"))[[1]],
    key = stringr::str_extract(param, "^[^=]*"),
    value = stringr::str_extract(param, "[^=]*$"),
  )
  room_param <- dplyr::filter(params, stringr::str_to_lower(value) == "room")
  dir_param <- dplyr::filter(params, stringr::str_to_lower(value) %in% c("enter", "exit"))
  
  values <- list(
    form_url = reg_exp_matches[,2],
    room_field_id = if (nrow(room_param) == 1) room_param$key else NA_character_,
    direction_field_id = if (nrow(dir_param) == 1) dir_param$key else NA_character_
  )
  
  return(values)
}

# generate qr code url
# generic function to generate qr code urls for any form and set of parameters
generate_qr_code_url <- function(
  form_url,
  form_parameters = c(),
  api_url = "http://api.qrserver.com/v1/create-qr-code",
  qr_code_size = "500x500"
) {
  form_url <- 
    sprintf(
      "%s/formResponse?%s&submit=SUBMIT", form_url,
      paste0(sprintf("%s=%s", names(form_parameters), purrr::map_chr(as.character(form_parameters), URLencode)), collapse = "&")
    )
  qr_code_url <- sprintf("%s/?data=%s&size=%s", api_url, URLencode(form_url, reserved = TRUE, repeated = TRUE), qr_code_size)
  return(qr_code_url)
}

# generate png file path
generate_png_file_path <- function(room, direction, pngs_dir = "pngs") {
  if (!dir.exists(pngs_dir)) dir.create(pngs_dir)
  return(file.path(pngs_dir, sprintf("%s_%s.png", room, direction)))
}

# generate docx save path
generate_docx_file_path <- function(room, direction, docx_dir = "docx") {
  if (!dir.exists(docx_dir)) dir.create(docx_dir)
  return(file.path(docx_dir, sprintf("%s-%s.docx", room, direction)))
}

# download qr code for room
# @param form_url the url of the google form (usually https://docs.google.com/forms/d/e/xxxxxxxx)
# @param room_field_id the ID of the room field in the google form, usually something like entry.12342356
# @param room the actual value for the room field
# @param direction_field_id the ID of the directtion field in the google form, usually something like entry.12342356
# @param direction usually either "Enter" or "Exit"
# @param ... additional parameters passed on to the generate_qr_code_url function
# @return png_path returns the png path
download_qr_code_for_room <- function(room, direction, form_url, room_field_id, direction_field_id, pngs_dir = "pngs", ...) {
  stopifnot(!missing(room) && !missing(direction) && !missing(form_url) && !missing(room_field_id) && !missing(direction_field_id))
  form_parameters <- c(room, direction) %>% setNames(c(room_field_id, direction_field_id))
  qr_code_url <- generate_qr_code_url(form_url = form_url, form_parameters = form_parameters, ...)
  png_path <- generate_png_file_path(room, direction)
  if (file.exists(png_path)) file.remove(png_path)
  download.file(qr_code_url, png_path)
  return(png_path)
}

# insert qr codes into a docx template
# @param template_path path to template docx
# @inheritParams download_qr_code_for_room
# @param data optional list of values to make available for code chunks inside the template doc (Room and Dir are automatically available)
# @param save whether to save the resulting docx
# @param save_path where to save it to if save=TRUE
# @return docx with barcode png tags evaluated, save with print(docx, target = filepath) or using save=TRUE
insert_qr_codes_into_doc <- function(template_path, room, direction, form_url, room_field_id, direction_field_id, data = list(), save = FALSE, save_path = generate_docx_file_path(room, direction)) {
  
  # template
  stopifnot(file.exists(template_path))
  doc <- officer::read_docx(template_path) 
  
  # Room and Direction info
  data$Room <- room
  data$Dir <- direction
  
  # function available to create barcode inside the document generation scope
  generate_qr_code <- function() {
    download_qr_code_for_room(
      room = room, direction = direction, 
      form_url = form_url, room_field_id = room_field_id, 
      direction_field_id = direction_field_id
    )
  }
  
  # regular expressions for r expression and png tags
  r_exp_regex <- "`r ([^`]*)`"
  png_exp_regex <- "`png(\\d+)x(\\d+) ([^`]*)`"
  
  # pull out styles
  styles <- officer::styles_info(doc)
  
  # get all text elements
  elements <- officer::docx_summary(doc) %>% 
    # add in the additional information
    dplyr::mutate(
      expr = ifelse(
        stringr::str_detect(text, r_exp_regex), 
        stringr::str_match_all(text, r_exp_regex), 
        list()),
      expr_value = purrr::map(expr, ~{
        if (length(.x) > 0) {
          expr_code <- .x[,2]
          values <- purrr::map_chr(expr_code, ~{
            tryCatch(rlang::parse_expr(.x) %>% rlang::eval_tidy(data = !!data) %>% as.character(),
                     error = function(e) {
                       warning(e$message, immediate. = TRUE, call. = FALSE)
                       "MISSING VALUE"
                     })
          }
          ) %>% unname()
          return(values)
        }
        else list()
      }),
      text_with_value = purrr::pmap_chr(
        list(text = text, expr = expr, expr_value = expr_value), 
        function(text, expr, expr_value) {
          if (length(expr) > 0) {
            full_expr <- expr[,1]
            for (i in 1:length(full_expr))
              text <- stringr::str_replace(text, fixed(full_expr[i]), expr_value[i])
            return(text)
          } else {
            return(text)
          }
        }),
      png = 
        ifelse(
          stringr::str_detect(text, png_exp_regex), 
          stringr::str_match(text, png_exp_regex)[,4], 
          NA_character_),
      png_width =  
        ifelse(
          stringr::str_detect(text, png_exp_regex), 
          stringr::str_match(text, png_exp_regex)[,2], 
          NA_character_) %>% as.numeric(),
      png_height =  
        ifelse(
          stringr::str_detect(text, png_exp_regex), 
          stringr::str_match(text, png_exp_regex)[,3], 
          NA_character_) %>% as.numeric(),
      png_path = purrr::map_chr(png, ~{
        if (!is.na(.x)) {
          rlang::parse_expr(.x) %>% rlang::eval_tidy(data = !!data) %>% as.character()
        }
        else NA_character_
      })   
    ) %>% 
    # add in styles
    dplyr::left_join(
      select(styles, style_type, style_name, style_id) %>% unique(), 
      by = c("content_type" = "style_type", "style_name")
    )
  
  # add interpreted expressions
  exprs <- dplyr::filter(elements, purrr::map_int(expr, length) > 0)
  if (nrow(exprs) > 0) {
    for (i in 1:nrow(exprs)) {
      doc <- 
        doc %>% 
        officer::cursor_reach(keyword = paste0("\\Q", exprs$text[i], "\\E")) %>% 
        officer::body_remove() %>% 
        officer::cursor_backward() %>% 
        officer::body_add_par(
          exprs$text_with_value[i], pos = "after", 
          style = if (is.na(exprs$style_name[i])) NULL else exprs$style_name[i]
        )
    }
  }
  
  # add pngs
  pngs <- dplyr::filter(elements, !is.na(png))
  if (nrow(pngs) > 0) {
    for (i in 1:nrow(pngs)) {
      doc <- 
        doc %>% 
        officer::cursor_reach(keyword = paste0("\\Q", pngs$text[i], "\\E")) %>% 
        officer::body_remove() %>% 
        officer::cursor_backward() %>% 
        officer::body_add_img(
          pngs$png_path[i],
          width = pngs$png_width[i], height = pngs$png_height[i],
          pos = "after", 
          style = if (is.na(exprs$style_name[i])) NULL else exprs$style_name[i]
        )
    }
  }
  
  # save?
  if (save) {
    if (file.exists(save_path)) file.remove(save_path)
    print(doc, target = save_path)
  }
  
  # return the doc
  return(invisible(doc))
}

# generate complete docx for one or more rooms
# 
# This is the main function to use to generate the QR codes docs.
#
# @inheritParams insert_qr_codes_into_doc
# @param data optional additional data frame, if provided must have a Room column and include all values in rooms
# @param directions - which directions to include, usually no need to change the default
# @export
generate_qr_codes_doc <- function(template_path, rooms, form_url, room_field_id, direction_field_id, data = tibble(Room = rooms), directions = c("Enter", "Exit"), save_path = sprintf("%s.docx", paste(rooms, collapse = "_"))) {
  
  # safety checks
  stopifnot(is.data.frame(data) && "Room" %in% names(data))
  stopifnot(all(rooms %in% data$Room))
  
  # generate individual room docs
  docs <- tibble::tibble(room = rooms) %>% 
    tidyr::crossing(tibble(direction = directions)) %>% 
    dplyr::mutate(
      save_path = purrr::map2_chr(room, direction, generate_docx_file_path) ,
      doc = purrr::map2(room, direction, ~{
        insert_qr_codes_into_doc(
          template_path = template_path,
          room = .x, 
          direction = .y,
          form_url = form_url, 
          room_field_id = room_field_id, 
          direction_field_id = direction_field_id,
          data = dplyr::filter(data, Room == .x)[1,] %>% as.list(),
          save = TRUE
        )
      })
    )
  
  # combine all docs
  doc <- docs$doc[[1]]
  if (nrow(docs) > 1) {
    for (i in 2:nrow(docs)) {
      doc <- doc %>% officer::cursor_end() %>% officer::body_add_docx(docs$save_path[i])
    }
  }
  if (file.exists(save_path)) file.remove(save_path)
  print(doc, target = save_path)
  
  return(invisible(docs$doc))
}