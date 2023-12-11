#'==============================================================================
#' R/3-amendment.R
#' 
#' - [1] 0-setup.R の実行チェック
#' - [2] 1-matching.R の実行チェック
#' - [3] 修正ファイルの読み込み
#' - [4] 修正後の資料作成
#'  
#' 実行コマンド
#' source("R/3-amendment.R", encoding = "utf-8")
#'
#'==============================================================================


#---- Is Setup Finished? -------------------------------------------------------
message("\n", "[3-amendment.R] Step 1. 0-setup.R の実行チェック", "\n")
#------------------------------------------------------------------------------#

if (!exists(".setup_finished", envir = .GlobalEnv) || isFALSE(.setup_finished)) {
  message("R/0-setup.R を実行しましたか？\n\n", 
          "     source(\"R/0-setup.R\", encoding = \"utf-8\")", "\n")
  stop("セットアップが完了していません。")
}

message("Done.")

#---- Is Matching Finished? -------------------------------------------------------
message("\n", "[3-amendment.R] Step 2. 1-matching.R の結果を読み込む。", "\n")
#------------------------------------------------------------------------------#

if (!file.exists(here::here(config$dir$out, "matching_data.rda"))) {
  message("R/1-matching.R を実行しましたか？\n\n", 
          "     source(\"R/1-matching.R\", encoding = \"utf-8\")", "\n")
  stop("マッチングデータが見つかりません。")
}

.matching_data <- readRDS(here::here(config$dir$out, "matching_data.rda"))
admin <- .matching_data$admin
config <- .matching_data$config
result <- .matching_data$result
util <- .matching_data$util

message("Done.")


#---- Create Documents -------------------------------------------------------
message("\n", "[3-amendment.R] Step 3. 修正ファイルの読み込み", "\n")
#------------------------------------------------------------------------------#

if (is.null(config$xls$change)) {
  stop("config.yml に修正ファイル (config > xls > change) の指定がありません。")
}

if (!file.exists(here::here(config$xls$change))) {
  stop("修正ファイルがありません。 ", config$xls$change)
}

.change_xls <- readxl::read_xlsx(config$xls$change)

if (any(is.na(match(c("Student", "Seminar", "Name"), names(.change_xls))))) {
  stop("修正ファイルの列名が (Student, Seminar, Name) とは異なります。")
}

message("Done.")


#---- Create Documents -------------------------------------------------------
message("\n", "[3-amendment.R] Step 4. 修正後の資料作成", "\n")
#------------------------------------------------------------------------------#

.matching_chg0 <- 
  result$match_table |> 
  rows_upsert(.change_xls |> select(Student, Seminar), by = "Student")

.matching_df_chg0 <-
  .matching_chg0 |> 
  left_join(admin$student, by = c("Student" = "ID")) 

.matching_df_chg1 <-
  .matching_df_chg0 |> 
  rows_patch(.change_xls |> select(Student, 氏名 = Name), by = "Student")

.matching_df_chg <-
  .matching_df_chg1 |> 
  mutate(ID_Name = paste(omueconMatch::last_n(Student, 3), 氏名)) |> 
  left_join(admin$faculty, by = c("Seminar" = "ID")) |> 
  rename(学生氏名 = 氏名, 教員氏名 = 教員名) |> 
  group_by(Seminar) |> 
  mutate(`通番` = row_number()) |> 
  ungroup() |> 
  select(!GPA)


##=============================================================================
##【修正後_マッチ掲示用.csv】
##=============================================================================
##
##> 学籍番号 配属ゼミ
##> <chr>    <chr>   
##> 1 STIDst1  教員A   
##> 2 STIDst2  教員A   
##> 3 STIDst3  教員C   
##> 4 STIDst4  教員B   
##> 5 STIDst5  教員B   
##> 6 STIDst6  教員C   
##>
.match_table_display_chg <- 
  .matching_df_chg |> 
  select(学籍番号 = Student, 配属ゼミ = 教員氏名)

readr::write_excel_csv(.match_table_display_chg, 
                       file = file.path(config$dir$out, "修正後_マッチ掲示用.csv"))

##=============================================================================
##【修正後_マッチMoodle用.html】
##=============================================================================
##
.template <- readLines(here::here("Template/result.Rmd"))
dir.create(.tdir <- tempfile(pattern = "dir-"))
.rendered <- whisker::whisker.render(
  .template, 
  data = list(result = paste(knitr::kable(.match_table_display_chg), collapse = "\n")))
writeLines(.rendered, file.path(.tdir, "修正後_マッチMoodle用.Rmd"))

juicedown::convert(file.path(.tdir, "修正後_マッチMoodle用.Rmd"), 
                   dir = config$dir$out, clip = FALSE)



##=============================================================================
##【修正後_マッチ名前入り.csv】
##【修正後_マッチ学籍番号.csv】
##=============================================================================

max_rows <- 
  .matching_df_chg |> 
  group_by(Seminar) |> 
  summarize(n = n()) |> 
  summarize(max = max(n)) |> 
  pull(max)

max_rows <- max(max_rows, config$slots)


report_with_id_chg <- tibble(通番 = seq_len(max_rows))
local({
  for (prof in admin$faculty$教員名) {
    report_with_id_chg[prof] <<- ""
  }
})

report_with_name_chg <- report_with_id_chg

local({
  for (prof in admin$faculty$教員名) {
    matched_students <- .matching_df_chg |> filter(教員氏名 == prof)
    st_names <- matched_students |> pull("ID_Name")
    st_id <- matched_students |> pull("Student")
    
    report_with_name_chg[seq_along(st_names), prof] <<- st_names
    report_with_id_chg[seq_along(st_id), prof] <<- st_id
  }
})

readr::write_excel_csv(report_with_name_chg, 
                       file = file.path(config$dir$out, "修正後_マッチ名前入り.csv"))
readr::write_excel_csv(report_with_id_chg, 
                       file = file.path(config$dir$out, "修正後_マッチ学籍番号.csv"))

##=============================================================================
##【修正後_空きゼミ一覧.csv】
##=============================================================================

.not_full_chg <-
  .matching_df_chg |> 
  group_by(教員氏名, Seminar) |> 
  summarize(空き枠 = config$slots - n(), .groups = "drop") |> 
  select(教員名 = 教員氏名, 空き枠) |> 
  filter(空き枠 > 0)

readr::write_excel_csv(.not_full_chg, 
                       file = file.path(config$dir$out, "修正後_空きゼミ一覧.csv"))

