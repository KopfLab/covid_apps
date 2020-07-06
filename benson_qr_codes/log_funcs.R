# data processing =====

# quick read
read_gs_sheet <- function(sheet_id, auth_token, ...) {
  
  # deauth
  googlesheets4::gs4_deauth()
  
  # read only authentication
  googlesheets4::gs4_auth(
    scopes = "https://www.googleapis.com/auth/spreadsheets.readonly",
    path = auth_token
  )
  
  if (googlesheets4::gs4_has_token()) {
    data <-
      tryCatch(
        googlesheets4::read_sheet(sheet_id, ...),
        error = function(e) {
          stop("google sheet access failed: ", e$message, call. = FALSE)
        }
      )
  } else {
    stop("google sheets authentication failed")
  }
}

download_data <- function(auth_token, rooms_sheet_id, logs_sheet_id) {
  rooms <- read_gs_sheet(sheet_id = rooms_sheet_id, auth_token = auth_token)
  raw_data <- read_gs_sheet(sheet_id = logs_sheet_id, auth_token = auth_token, col_types = "TcccTTc")
  raw_data %>% 
    left_join(rooms, by = "Room")
}

process_raw_data <- function(raw_data) {
  # process raw data
  raw_data %>% 
    # avoid troubles with non-specified rooms in the logs
    filter(!is.na(Floor)) %>% 
    select(datetime = Timestamp, user = `Email Address`, room = Room, dir = Direction, floor = Floor, notes = Notes) %>% 
    # account for BESC "room"
    mutate(
      room_id = 
        case_when(
          stringr::str_detect(room, "BESC") ~ stringr::str_remove(room, "(East|West|North|South)$"),
          TRUE ~ room,
        ),
      datetime = lubridate::force_tz(datetime, "America/Denver")
    ) %>% 
    # determine room occupancy
    arrange(user, room_id, datetime) %>% 
    group_by(user, floor, room_id) %>% 
    mutate(
      enter_grp = cumsum(dir == "Enter"),
      exit_grp = cumsum(dir == "Exit"),
      has_enter = dir == "Exit" & c(FALSE, head(dir, -1) == "Enter"),
      has_exit = dir == "Enter" & c(tail(dir, -1) == "Exit", FALSE),
      room_grp = ifelse(has_enter | has_exit, cumsum(has_exit), 0),
    ) %>% 
    ungroup() %>% 
    # create factors
    arrange(floor, room_id) %>% 
    mutate(
      user_id = stringr::str_remove(user, "@colorado.edu"), 
      floor_factor = factor(floor) %>% forcats::fct_inorder() %>% forcats::fct_relevel("basement"), 
      room_factor = factor(room_id) %>% forcats::fct_inorder(),
      day = format(datetime, "%b %d")
    )
}

filter_data <- function(data, users, start_date = NULL, end_date = NULL) {
  data %>% 
    filter(user_id %in% users) %>% 
    { if (!is.null(start_date)) filter(., datetime > start_date) else . } %>% 
    { if (!is.null(end_date)) filter(., datetime < end_date) else . } 
}

# data preparation =====

prepare_plotting_data <- function(data) {
  left_join(
    filter(data, dir == "Enter") %>% 
      rename(start_datetime = datetime) %>% 
      select(-enter_grp, -exit_grp, -has_enter, -has_exit),
    filter(data, dir == "Exit", room_grp > 0) %>% 
      select(user_id, room_id, room_grp, end_datetime = datetime) ,
    by = c("user_id", "room_id", "room_grp")
  ) %>% 
    bind_rows(
      filter(data, dir == "Exit", room_grp == 0) %>% 
        rename(end_datetime = datetime) %>% 
        select(-enter_grp, -exit_grp, -has_enter, -has_exit)
    ) %>% 
    mutate(
      datetime = case_when( 
        !is.na(start_datetime) ~ start_datetime, 
        !is.na(end_datetime) ~ end_datetime,
        TRUE ~ lubridate::as_datetime(NA_real_)
      )
    ) %>% 
    arrange(user_id, room_id, datetime)
}

process_for_show_by <- function(data, show_by = c("BUILDING", "FLOOR", "ROOM"), include_unmatched = show_by == "ROOM") {
  
  # total y range
  y_range = 0.8
  
  # note: this probably won't work b/c how to show multiple people in the same room?
  plot_data <- data %>% 
    filter(show_by == "BUILDING" | room_id != "BESC") %>% 
    # whether to include unmatched
    { if (!include_unmatched) filter(., room_grp > 0) else . } %>% 
    arrange(datetime) %>% 
    mutate(
      day_factor = factor(day) %>% forcats::fct_inorder(),
      user_factor = factor(user_id) %>% forcats::fct_inorder(),
      room_factor = room_factor %>% forcats::fct_drop(),
      floor_factor = floor_factor %>% forcats::fct_drop(),
      room_nr = as.numeric(room_factor),
      floor_nr = as.numeric(floor_factor)
    ) 
  
  # what to filter by
  if (show_by == "ROOM") {
    plot_data <- plot_data %>% 
      mutate(
        y_breaks = room_nr,
        y_start = room_nr - y_range/2 + (as.numeric(user_factor) - 1) * y_range / length(unique(user_id)),
        y_end = y_start + y_range / length(unique(user_id)),
        y = (y_start + y_end)/2,
        y_labels = room_factor
      )
  } else if (show_by == "FLOOR") {
    plot_data <- plot_data %>% 
      mutate(
        y_breaks = floor_nr,
        y_start = floor_nr - y_range/2 + (as.numeric(user_factor) - 1) * y_range / length(unique(user_id)),
        y_end = y_start + y_range / length(unique(user_id)),
        y = (y_start + y_end)/2,
        y_labels = floor_factor
      ) 
  } else if (show_by == "BUILDING") {
    plot_data <- plot_data %>% 
      mutate(
        y_breaks = 0,
        y_start = -y_range/2 + (as.numeric(user_factor) - 1) * y_range / length(unique(user_id)),
        y_end = y_start + y_range / length(unique(user_id)),
        y = (y_start + y_end)/2,
        y_labels = factor("BESC")
      ) 
  } else {
    stop("don't know show_by ", show_by, call. = FALSE)
  }
  return(plot_data)
}

# visualization ======

plot_logs <- function(plot_data) {
  
  # safety checks
  if (nrow(plot_data) == 0)
    return(NULL)
  
  # setup plot
  p <- plot_data %>% 
    ggplot() +
    aes(y = y, color = user_id)
  
  # whether to show unmatched markers
  if (any(plot_data$room_grp == 0)) {
    p <- p + geom_point(
      data = function(df) 
        filter(df, room_grp == 0) %>% 
        mutate(dir = paste("Unmatched", dir)),
      mapping = aes(x = datetime, shape = dir),
      size = 4
    )
  }

  # main plot
  p +
    geom_rect(
      data = function(df) filter(df, room_grp > 0),
      mapping = aes(
        xmin = start_datetime, xmax = end_datetime,
        ymin = y_start, ymax = y_end, 
        fill = user_id, color = NULL
      ),
      size = 3
    ) + 
    scale_shape_manual(values = c(2, 6)) +
    scale_color_brewer(palette = "Set1") +
    scale_fill_brewer(palette = "Set1") + 
    scale_x_datetime(date_labels = "%H:%M") + 
    facet_wrap(~day_factor, scales = "free_x") +
    scale_y_continuous(
      breaks = sort(unique(plot_data$y_breaks)), labels = levels(plot_data$y_labels)
    ) +
    theme_bw() +
    theme(
      panel.grid = element_blank(), 
      panel.grid.major.y = element_line(color = "gray", linetype = 2),
      panel.grid.minor.y = element_line(color = "black"),
      text = element_text(size = 20), 
      axis.text.x = element_text(angle = 90, hjust = 0, vjust = 0.5),
      axis.ticks.y = element_blank()
    ) +
    guides(color = FALSE, fill = FALSE) +
    labs(x = NULL, y = NULL, shape = NULL)
}