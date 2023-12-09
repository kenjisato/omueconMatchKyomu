#'==============================================================================
#' R/1-matching.R
#' 
#' - [1] 0-setup.R の実行チェック
#' - [2] 選好表の生成
#' - [3] マッチング
#' - [4] データの保存
#'  
#' 実行コマンド
#' source("R/1-matching.R", encoding = "utf-8")
#'
#'==============================================================================

#---- Is Setup Finished? -------------------------------------------------------
message("\n", "[1-matching.R] Step 1. 0-setup.R の実行チェック", "\n")
#------------------------------------------------------------------------------#

if (!exists(".setup_finished", envir = .GlobalEnv) || isFALSE(.setup_finished)) {
  message("R/0-setup.R を実行しましたか？\n\n", 
          "     source(\"R/0-setup.R\", encoding = \"utf-8\")", "\n")
  stop("セットアップが完了していません。")
}

message("Done.")

#---- Create Utility Table -----------------------------------------------------
message("\n", "[1-matching.R] Step 2. 選好表の生成", "\n")
#------------------------------------------------------------------------------#

# 教務事務管理ファイルの読み込み
admin <- list()
admin$student <- readxl::read_excel(here::here(config$xls$student), sheet = 1)
admin$faculty <- readxl::read_excel(here::here(config$xls$faculty), sheet = 1)

# 選好表の作成
util <- omueconMatch::matching_utils(
    student_list = admin$student,
    faculty_list = admin$faculty,
    dir_student = here::here(config$dir$student),
    dir_faculty = here::here(config$dir$faculty),
    nc = config$nc,
    seed = config$seed
) 

message("Done.")

#---- Match --------------------------------------------------------------------
message("\n", "[1-matching.R] Step 3. マッチング", "\n")
#------------------------------------------------------------------------------#

result <- omueconMatch::matching_compute(util, slots = config$slots)

message("Done.")

#---- Save Data ----------------------------------------------------------------
message("\n", "[1-matching.R] Step 4. データの保存", "\n")
#------------------------------------------------------------------------------#

.matching_result <- list(
  config = config,
  admin = admin,
  util = util,
  result = result
)
saveRDS(.matching_result, 
        here::here(config$dir$out, "matching_data.rda"))

message("Done")
