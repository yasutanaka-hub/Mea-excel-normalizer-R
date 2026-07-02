# ==== Multi-row extractor for Windows ====

# ===== 必要パッケージ =====
required_packages <- c(
  "readxl",
  "dplyr",
  "tidyr",
  "tibble",
  "purrr",
  "stringr",
  "writexl"
)

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
}

library(readxl)
library(dplyr)
library(tidyr)
library(tibble)
library(purrr)
library(stringr)
library(writexl)

# ===== 設定 =====

# Windows用：Desktop/R/excel_files を入力フォルダにする
folder_path <- file.path(Sys.getenv("USERPROFILE"), "Desktop", "R", "excel_files")

# Windows用：Desktop/R/excel_output_files を出力フォルダにする
out_dir <- file.path(Sys.getenv("USERPROFILE"), "Desktop", "R", "excel_output_files")

if (!dir.exists(out_dir)) {
  dir.create(out_dir, recursive = TRUE)
}

if (!dir.exists(folder_path)) {
  stop("入力フォルダが見つかりません: ", folder_path)
}

# .xlsx 取得後、~$ と ._ で始まる一時/隠しファイルを除外
file_list <- list.files(path = folder_path, pattern = "\\.xlsx$", full.names = TRUE)
file_list <- file_list[!grepl("(^~\\$)|(^\\._)", basename(file_list))]

if (length(file_list) == 0) {
  stop("xlsxファイルが見つかりません: ", folder_path)
}

# 列レンジ
col_start <- "B"
col_end   <- "AW"

# .bat から行番号を受け取る
# 指定がなければデフォルト行を使う
args <- commandArgs(trailingOnly = TRUE)

if (length(args) > 0) {
  target_rows <- suppressWarnings(as.integer(args))
  target_rows <- target_rows[!is.na(target_rows)]
} else {
  target_rows <- c(124, 129, 131, 133, 135, 137, 139)
}

if (length(target_rows) == 0) {
  stop("target_rows が指定されていません。")
}

# 固定のシート名（NULLで1枚目）
sheet_name <- NULL

# 列名にベースラベル(A列)を含めるか
use_base_in_colnames <- TRUE

# ===== Utils =====

make_range <- function(row, start_col = "B", end_col = "AW") {
  paste0(start_col, row, ":", end_col, row)
}

col_to_index <- function(col_letters) {
  s <- strsplit(toupper(col_letters), "")[[1]]
  sum((match(s, LETTERS)) * 26^(rev(seq_along(s)) - 1))
}

fixed_ncols <- col_to_index(col_end) - col_to_index(col_start) + 1

safe_read_cell_text <- function(file, addr, sheet = NULL) {
  tryCatch(
    suppressMessages(
      read_excel(
        file,
        sheet = sheet,
        range = addr,
        col_names = FALSE,
        col_types = "text"
      )
    )[[1]],
    error = function(e) NA_character_
  )
}

safe_read_range_text <- function(file, rng, sheet = NULL) {
  tryCatch(
    suppressMessages(
      read_excel(
        file,
        sheet = sheet,
        range = rng,
        col_names = FALSE,
        col_types = "text"
      )
    ),
    error = function(e) NULL
  )
}

safe_read_range_any <- function(file, rng, sheet = NULL) {
  tryCatch(
    suppressMessages(
      read_excel(
        file,
        sheet = sheet,
        range = rng,
        col_names = FALSE
      )
    ),
    error = function(e) NULL
  )
}

build_sub_labels <- function(file, start_col, end_col, k, sheet = NULL) {
  base <- paste0("Day", seq_len(k))

  df <- safe_read_range_text(
    file,
    paste0(start_col, "122:", end_col, "122"),
    sheet = sheet
  )

  if (is.null(df)) return(base)

  tmp <- trimws(unlist(df[1, , drop = TRUE]))

  if (length(tmp) == 0) return(base)

  if (length(tmp) >= k) {
    tmp <- tmp[seq_len(k)]
  } else {
    tmp <- c(tmp, rep(NA_character_, k - length(tmp)))
  }

  idx <- which(is.na(tmp) | tmp == "")
  if (length(idx) > 0) {
    tmp[idx] <- paste0("Day", idx)
  }

  tmp
}

read_values_or_na_row <- function(file, row_num, k, start_col, end_col, sheet = NULL) {
  rng <- make_range(row_num, start_col, end_col)

  vals <- safe_read_range_any(file, rng, sheet = sheet)

  if (is.null(vals) || ncol(vals) == 0 || nrow(vals) == 0) {
    mat <- matrix(
      NA_real_,
      nrow = 1,
      ncol = k,
      dimnames = list(NULL, paste0("V", seq_len(k)))
    )

    return(as_tibble(as.data.frame(mat), .name_repair = "minimal"))
  }

  vals <- vals %>%
    mutate(across(everything(), ~ suppressWarnings(as.numeric(.))))

  if (ncol(vals) < k) {
    pad <- matrix(
      NA_real_,
      nrow = nrow(vals),
      ncol = k - ncol(vals),
      dimnames = list(NULL, paste0("Pad", seq_len(k - ncol(vals))))
    )

    vals <- bind_cols(vals, as_tibble(pad, .name_repair = "minimal"))

  } else if (ncol(vals) > k) {
    vals <- vals[, seq_len(k), drop = FALSE]
  }

  vals
}

# ===== 1行ぶんを全ファイルで処理 =====

process_one_row <- function(row_index) {
  results <- vector("list", length(file_list))

  for (i in seq_along(file_list)) {
    file <- file_list[i]

    message("Processing: ", basename(file), " @ row ", row_index)

    # ベースラベル（A{row_index}）。空は "Label"
    base_raw <- safe_read_cell_text(
      file,
      paste0("A", row_index),
      sheet = sheet_name
    )

    base_label <- ifelse(
      is.na(base_raw) || base_raw == "",
      "Label",
      base_raw
    )

    # サブラベル（固定長）
    sub_labels <- build_sub_labels(
      file,
      col_start,
      col_end,
      fixed_ncols,
      sheet = sheet_name
    )

    # 値（固定長、空でもNA行）
    values <- read_values_or_na_row(
      file,
      row_index,
      fixed_ncols,
      col_start,
      col_end,
      sheet = sheet_name
    )

    # 列名
    if (use_base_in_colnames) {
      new_names <- paste0(base_label, "_", sub_labels)
    } else {
      new_names <- sub_labels
    }

    # 列名の重複を避ける
    colnames(values) <- make.unique(new_names, sep = "_")

    # 識別列
    out <- values %>%
      mutate(
        File = tools::file_path_sans_ext(basename(file)),
        Row  = row_index,
        .before = 1
      )

    results[[i]] <- out
  }

  bind_rows(results)
}

# ===== 行ごとにシートを作成 =====

sheet_list <- lapply(target_rows, process_one_row)
names(sheet_list) <- paste0("Row", target_rows)

# ===== 出力 =====

timestamp <- format(Sys.time(), "%Y%m%d_%H%M")

row_label <- paste(target_rows, collapse = "-")

# ファイル名が長くなりすぎる場合に備えて短縮
if (nchar(row_label) > 80) {
  row_label <- paste0(length(target_rows), "rows")
}

output_filename <- paste0(
  "output_rows_",
  row_label,
  "_",
  timestamp,
  ".xlsx"
)

output_path <- file.path(out_dir, output_filename)

write_xlsx(sheet_list, path = output_path)

message("✅ Saved to: ", output_path)
