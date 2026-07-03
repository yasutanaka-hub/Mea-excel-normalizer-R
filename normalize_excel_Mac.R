# 任意の行(target_row)を任意の行(denom_row)で標準化し、3シートで出力
# ==== Normalize Excel rows for Mac ====

# ===== 必要パッケージ =====
required_packages <- c("readxl", "dplyr", "tibble", "writexl")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
}

library(readxl)
library(dplyr)
library(tibble)
library(writexl)

# ===== 設定 =====

# Mac用：Desktop/R/excel_files を入力フォルダにする
base_dir <- file.path(path.expand("~"), "Desktop", "R")

folder_path <- file.path(base_dir, "excel_files")
output_folder <- file.path(base_dir, "excel_output_files")

# .command から行番号を受け取る
args <- commandArgs(trailingOnly = TRUE)

target_row <- if (length(args) >= 1) as.integer(args[1]) else 129
denom_row  <- if (length(args) >= 2) as.integer(args[2]) else 130

if (is.na(target_row) || is.na(denom_row)) {
  stop("target_row または denom_row が数値として読み取れません。")
}

# 列レンジ
col_start <- "B"
col_end   <- "AW"

# 出力フォルダがなければ作成
if (!dir.exists(output_folder)) {
  dir.create(output_folder, recursive = TRUE)
}

# 入力フォルダが存在しない場合は停止
if (!dir.exists(folder_path)) {
  stop("入力フォルダが見つかりません: ", folder_path)
}

# Excelファイル一覧
# ~$ と ._ で始まる一時/隠しファイルを除外
file_list <- list.files(
  path = folder_path,
  pattern = "\\.xlsx$",
  full.names = TRUE
)

file_list <- file_list[!grepl("(^~\\$)|(^\\._)", basename(file_list))]

if (length(file_list) == 0) {
  stop("xlsxファイルが見つかりません: ", folder_path)
}

# ユーティリティ：列レンジを作る
make_range <- function(row, start_col = "B", end_col = "AW") {
  paste0(start_col, row, ":", end_col, row)
}

# 収集用リスト
norm_list <- list()
tgt_list  <- list()
den_list  <- list()

for (file in file_list) {
  message("Processing: ", basename(file))

  target_label <- tryCatch(
    suppressMessages(
      read_excel(
        file,
        range = paste0("A", target_row),
        col_names = FALSE,
        col_types = "text"
      )[[1]]
    ),
    error = function(e) NA
  )

  if (is.na(target_label) || target_label == "") {
    target_label <- "Target"
  }

  denom_label <- tryCatch(
    suppressMessages(
      read_excel(
        file,
        range = paste0("A", denom_row),
        col_names = FALSE,
        col_types = "text"
      )[[1]]
    ),
    error = function(e) NA
  )

  if (is.na(denom_label) || denom_label == "") {
    denom_label <- "Denominator"
  }

  sub_df <- tryCatch(
    suppressMessages(
      read_excel(
        file,
        range = paste0(col_start, "122:", col_end, "122"),
        col_names = FALSE,
        col_types = "text"
      )
    ),
    error = function(e) NULL
  )

  if (is.null(sub_df)) {
    next
  }

  sub_labels <- trimws(unlist(sub_df[1, , drop = TRUE]))

  sub_labels <- ifelse(
    is.na(sub_labels) | sub_labels == "",
    paste0("Day", seq_along(sub_labels)),
    sub_labels
  )

  tgt <- tryCatch(
    suppressMessages(
      read_excel(
        file,
        range = make_range(target_row, col_start, col_end),
        col_names = FALSE
      )
    ),
    error = function(e) NULL
  )

  den <- tryCatch(
    suppressMessages(
      read_excel(
        file,
        range = make_range(denom_row, col_start, col_end),
        col_names = FALSE
      )
    ),
    error = function(e) NULL
  )

  if (is.null(tgt) || is.null(den)) {
    next
  }

  tgt <- tgt %>%
    mutate(across(everything(), ~ suppressWarnings(as.numeric(.))))

  den <- den %>%
    mutate(across(everything(), ~ suppressWarnings(as.numeric(.))))

  target_cols <- length(sub_labels)

  pad_to <- function(df, k) {
    if (ncol(df) >= k) {
      return(df[, 1:k, drop = FALSE])
    }

    pad <- as_tibble(
      matrix(
        NA_real_,
        nrow = nrow(df),
        ncol = k - ncol(df)
      ),
      .name_repair = "minimal"
    )

    bind_cols(df, pad)
  }

  tgt <- pad_to(tgt, target_cols)
  den <- pad_to(den, target_cols)

  tgt_mat <- as.matrix(tgt)
  den_mat <- as.matrix(den)

  # ====== 1) 標準化 ======
  norm_mat <- tgt_mat / den_mat

  na_mask    <- is.na(tgt_mat) | is.na(den_mat)
  denom_zero <- (!is.na(den_mat)) & den_mat == 0

  norm_mat[na_mask] <- NA
  norm_mat[denom_zero & tgt_mat == 0] <- 0
  norm_mat[denom_zero & tgt_mat != 0] <- NA

  norm <- as_tibble(norm_mat, .name_repair = "minimal")
  colnames(norm) <- make.unique(
    paste0(target_label, " (per ", denom_label, ")_", sub_labels),
    sep = "_"
  )

  norm <- norm %>%
    mutate(File = tools::file_path_sans_ext(basename(file)), .before = 1)

  norm_list[[length(norm_list) + 1]] <- norm

  # ====== 2) target 生データ ======
  tgt_out <- as_tibble(tgt_mat, .name_repair = "minimal")
  colnames(tgt_out) <- make.unique(
    paste0(target_label, "_", sub_labels),
    sep = "_"
  )

  tgt_out <- tgt_out %>%
    mutate(File = tools::file_path_sans_ext(basename(file)), .before = 1)

  tgt_list[[length(tgt_list) + 1]] <- tgt_out

  # ====== 3) denom 生データ ======
  den_out <- as_tibble(den_mat, .name_repair = "minimal")
  colnames(den_out) <- make.unique(
    paste0(denom_label, "_", sub_labels),
    sep = "_"
  )

  den_out <- den_out %>%
    mutate(File = tools::file_path_sans_ext(basename(file)), .before = 1)

  den_list[[length(den_list) + 1]] <- den_out
}

if (length(norm_list) == 0) {
  stop("処理できるデータがありませんでした。Excelの行番号・列範囲を確認してください。")
}

final_norm <- bind_rows(norm_list)
final_tgt  <- bind_rows(tgt_list)
final_den  <- bind_rows(den_list)

timestamp <- format(Sys.time(), "%Y%m%d_%H%M")

output_filename <- paste0(
  "output_norm_row",
  target_row,
  "_by_row",
  denom_row,
  "_",
  timestamp,
  ".xlsx"
)

output_path <- file.path(output_folder, output_filename)

write_xlsx(
  x = list(
    Sheet1 = final_norm,
    Sheet2 = final_tgt,
    Sheet3 = final_den
  ),
  path = output_path
)

message("✅ Saved to: ", output_path)
