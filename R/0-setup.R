#'==============================================================================
#' R/0-setup.R
#'  
#' - [1] 必要なパッケージのインストール
#' - [2] 設定ファイルの読み込み
#' - [3] 学生側評価のファイル名修正
#'  
#' 実行コマンド
#' source("R/0-setup.R", encoding = "utf-8")
#'
#'==============================================================================

.setup_finished <- FALSE

#---- Required Packages --------------------------------------------------------
message("\n", "[0-setup.R] Step 1. 必要なパッケージのインストール", "\n")
#------------------------------------------------------------------------------#

if (!require("remotes"))
  install.packages("remotes", repos = "https://cloud.r-project.org")

remotes::install_github("kenjisato/omueconMatch", upgrade = "never")

library(dplyr, warn.conflicts = FALSE)
library(ggplot2)
library(omueconMatch)


#---- Config -------------------------------------------------------------------
message("\n", "[0-setup.R] Step 2. 設定ファイルの読み込み", "\n")
#'-----------------------------------------------------------------------------#

# 端末のロカールが SJIS かもしれないので、UTF-8に変更
if (.Platform$OS.type == "windows") {
  Sys.setlocale(locale = "Japanese_Japan.utf8") 
}

# 設定ファイルの指定
# "default", "test1" or "test2"
#
.config <- "test1"

# 設定ファイル読み込み
config <- config::get(config = .config)

.spf <- function(s) paste0("- ", s, " : ")
message("[設定] \n",
        .spf("教員リスト"), config$xls$faculty, "\n",
        .spf("学生リスト"), config$xls$student, "\n",
        .spf("教員側評価"), config$dir$faculty, "\n",
        .spf("学生側評価"), config$dir$student, "\n",
        .spf("出力先　　"), config$dir$out, "\n")

#=============================================================================


#---- Rename Student Files------------------------------------------------------
message("\n", "[0-setup.R] Step 3. 学生側評価のファイル名修正", "\n")
#------------------------------------------------------------------------------#

# 学生ファイルの名前が長すぎて読み込めないので修正する。
.student_files <- list.files(here::here(config$dir$student), "\\.xlsx$")
.student_before <- .student_files[nchar(.student_files) > (config$nc + 5)]

if (length(.student_before) != 0) {
  .student_after <- paste0(substr(.student_before, 1, config$nc), ".xlsx")
  file.rename(
    here::here(config$dir$student, .student_before),
    here::here(config$dir$student, .student_after)
  )
  message(length(.student_before), "個のファイルを修正しました。", "\n")
} else {
  message("...Skipped....", "\n")
}


# Is setup finished?
.setup_finished <- TRUE
