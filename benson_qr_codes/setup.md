# Setup

adapted from https://www.youtube.com/watch?v=vISRn5qFrkM

 - go to https://console.developers.google.com/
 - switch to correct google account and note in the URL which `authuser` it is (e.g. 3)
 - click on `select a project` and make `NEW PROJECT` and set up new project with an informative name
 - go to https://console.developers.google.com/apis/dashboard?authuser=3 (change `authuser` as needed)
 - select the newly created project from the dropdown
 - go to `Library` in left hand menu, search for `google sheets api` and click `ENABLE`
 - go to https://console.developers.google.com/apis/credentials?authuser=3 (change `authuser` as needed)
 - make sure the correct project is selected
 - click on `CREATE CREDENTIALS`
 - select the `Google Sheets API`
 - then `Web Server` (that's where we're calling from in R)
 - then `Application Data`
 - then `No, I'm not using` for the APP engine
 - then `What credentials do I need?` to get to the next screen
 - create a service acount with e.g. name `editor` and select the role `Editor` to provide this account edit access to the google sheets (you can use `Viewer` only if read-only access is all that is needed)
 - select Key type `JSON` and `CONTINUE` to download the access token in json format (e.g. `client_secret.json`)
 - open the `.json` file and share the spreadsheet(s) you want to access with the `client_email` address
 - --> use `googlesheets4` to access the sheet based on its sheet ID (check URL), basic example
 
```
# authentication
googlesheets4::gs4_auth(path = "editor_client_secret.json")

# sheet ID
sheet_id <- "1jfj2u5hm-vIrq1oXiNEGgvEQKpsbovvCZ2iD6uZ3zMM"

if (googlesheets4::gs4_has_token()) {
  data <- 
    tryCatch(
      googlesheets4::read_sheet(sheet_id),
      error = function(e) {
        stop("google sheet access failed: ", e$message, call. = FALSE)
      }
    )
} else {
  stop("google sheets authentication failed")
}
```
