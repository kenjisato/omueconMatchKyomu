# install.packages("remotes", repos = "https://cloud.r-project.org")
# remotes::install_github("kenjisato/omueconMatch")
# source("R/matching.R", encoding = "utf-8")

library(dplyr, warn.conflicts = FALSE)
library(ggplot2)
library(omueconMatch)

#################################################
#
#  本番のマッチングの際には下のように書き換える
#  .config <- "default"
#
# "default", "test1" or "test2"
#
.config <- "default"

#################################################

# 設定ファイル（config.yml）の読み込み
config <- config::get(config = .config)


# File rename
# 学生ファイルの名前が長すぎて読み込めないので修正する。
#
for (f in list.files(here::here(config$dir$student))) {
  if (nchar(f) <= config$nc + 5) next()
  
  file.rename(
    here::here(config$dir$student, f),
    here::here(config$dir$student, paste0(substr(f, 1, config$nc), ".xlsx"))
  )
}


# 教務事務管理ファイルの読み込み
admin_student_list <- readxl::read_excel(here::here(config$xls$student), sheet = 1)
admin_faculty_list <- readxl::read_excel(here::here(config$xls$faculty), sheet = 1)

# 選好表の作成
util <- omueconMatch::matching_utils(
    student_list = admin_student_list,
    faculty_list = admin_faculty_list,
    dir_student = here::here(config$dir$student),
    dir_faculty = here::here(config$dir$faculty),
    nc = config$nc,
    seed = config$seed
) 

# マッチングの計算
result <- omueconMatch::matching_compute(util, slots = config$slots)

# 掲示物の作成
matching_result <-
  result$match_table |> 
  left_join(admin_student_list, by = c("Student" = "ID")) |> 
  mutate(ID_Name = paste(omueconMatch::last_n(Student, 3), 氏名)) |> 
  left_join(admin_faculty_list, by = c("Seminar" = "ID")) |> 
  rename(学生氏名 = 氏名, 教員氏名 = 教員名) |> 
  group_by(Seminar) |> 
  mutate(`通番` = row_number()) |> 
  ungroup() |> 
  select(!GPA)

report_with_id <- tibble(通番 = seq_len(config$slots))
for (prof in admin_faculty_list$教員名) {
  report_with_id[prof] <- ""
}
report_with_name <- report_with_id 
  
##=============================================================================
##【マッチ掲示用.csv】
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
match_table_display <- 
  matching_result |> 
  select(学籍番号 = Student, 配属ゼミ = 教員氏名)

readr::write_excel_csv(match_table_display, 
                       file = file.path(config$dir$out, "マッチ掲示用.csv"))

##=============================================================================
##【マッチ掲示用.html】
##=============================================================================
##
template <- readLines(here::here("Template/result.Rmd"))
dir.create(tdir <- tempfile(pattern = "dir-"))
rendered <- whisker::whisker.render(template, 
    data = list(result = paste(knitr::kable(match_table_display), collapse = "\n")))
writeLines(rendered, file.path(tdir, "マッチMoodle用.Rmd"))

juicedown::convert(file.path(tdir, "マッチMoodle用.Rmd"), 
                   dir = config$dir$out, clip = FALSE)

##=============================================================================
##【マッチ名前入り.csv】
##【マッチ学籍番号.csv】
##=============================================================================


for (prof in admin_faculty_list$教員名) {
  matched_students <- matching_result |> filter(教員氏名 == prof)
  st_names <- matched_students |> pull("ID_Name")
  st_id <- matched_students |> pull("Student")
  
  report_with_name[seq_along(st_names), prof] <- st_names
  report_with_id[seq_along(st_id), prof] <- st_id
}


readr::write_excel_csv(report_with_name, 
                       file = file.path(config$dir$out, "マッチ名前入り.csv"))

readr::write_excel_csv(report_with_id, 
                       file = file.path(config$dir$out, "マッチ学籍番号.csv"))

##=============================================================================
##【空きゼミ一覧.csv】
##=============================================================================
##

not_full <- 
  tibble(担当教員 = rownames(util$Student)[result$unmatched.colleges]) |> 
  group_by(担当教員) |> 
  summarize(空き枠 = n()) |> 
  ungroup() |> 
  left_join(admin_faculty_list, by = c("担当教員" = "ID")) |> 
  select(教員名, 空き枠)

readr::write_excel_csv(not_full, 
                       file = file.path(config$dir$out, "空きゼミ一覧.csv"))


##=============================================================================
##【目検チェック用】チェック_学生側評価.csv, チェック_教員側評価.csv
##=============================================================================

util_st_rounded <- 100 - round(util$Student)
write.csv(util_st_rounded, 
          file = file.path(config$dir$out, "チェック_学生側評価.csv"))

write.csv(round(util$Faculty, 2), 
          file = file.path(config$dir$out, "チェック_教員側評価.csv"))

##=============================================================================
##【統計情報】チェック_マッチ順位.csv
##=============================================================================

match_stat <- result$match_table
match_stat$Rank <- 0L
for (i in seq_len(nrow(match_stat))) {
  student <- match_stat$Student[i]
  seminar <- match_stat$Seminar[i]
  match_stat$Rank[i] <- util_st_rounded[seminar, student]
}

readr::write_excel_csv(
  match_stat, 
  file = file.path(config$dir$out, "チェック_マッチ順位.csv")
)

##=============================================================================
##【統計情報】チェック_マッチ順位.png
##=============================================================================
p <- ggplot(match_stat, aes(x = Rank)) + 
  geom_bar(fill = "skyblue") + 
  theme_linedraw()
ggsave(file.path(config$dir$out, "チェック_マッチ順位.png"), p, dpi = 300, 
       width = 800, height = 480, units = "px")

##=============================================================================
##【統計情報】マッチ順位on順位付けの数.png
##=============================================================================

# https://joshuacook.netlify.app/post/integer-values-ggplot-axis/
integer_breaks <- function(n = 5, ...) {
  fxn <- function(x) {
    breaks <- floor(pretty(x, n, ...))
    names(breaks) <- attr(breaks, "labels")
    breaks
  }
  return(fxn)
}

num_ranked_profs <- 
  tibble::as_tibble(colSums(util_st_rounded < 100), 
                    rownames = "Student") |> 
  rename(NumRanked = value)

match_stat2 <- match_stat |> 
  left_join(num_ranked_profs, by = "Student")

q <- ggplot(match_stat2, aes(x = NumRanked, y = Rank)) + 
  geom_jitter(width = 0.05, height = 0.05, size = 3, alpha = 0.4) + 
  xlab("Number of seminars ranked") + 
  ylab("Ranking given to \nassigned seminar") +
  scale_x_continuous(breaks = integer_breaks()) +
  scale_y_reverse(breaks = integer_breaks()) +
  theme_linedraw()
  
ggsave(file.path(config$dir$out, "チェック_マッチ順位onランク数.png"), q, dpi = 300, 
       width = 800, height = 600, units = "px")
