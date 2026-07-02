# CSVファイルをExcel形式（.xlsx）に一括変換するWindows版スクリプト

# ===== 必要パッケージ =====
required_packages <- c("readr", "writexl", "tibble")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
}

library(readr)
library(writexl)
library(tibble)

# ===== 設定 =====

# Windows用：Desktop/R/csv_files を入力フォルダにする
input_dir <- file.path(Sys.getenv("USERPROFILE"), "Desktop", "R", "csv_files")

# Windows用：Desktop/R/excel_conv_output_files を出力フォルダにする
output_dir <- file.path(Sys.getenv("USERPROFILE"), "Desktop", "R", "excel_conv_output_files")

# 出力フォルダがなければ作成
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

# 入力フォルダが存在しない場合は停止
if (!dir.exists(input_dir)) {
  stop("入力フォルダが見つかりません: ", input_dir)
}

# CSVファイル一覧
csv_files <- list.files(
  input_dir,
  pattern = "\\.[Cc][Ss][Vv]$",
  full.names = TRUE
)

if (length(csv_files) == 0) {
  stop(sprintf("CSVが見つかりませんでした: %s", input_dir))
}

# ===== 共通ユーティリティ =====

make_xlsx_path <- function(csv_path) {
  fn <- tools::file_path_sans_ext(basename(csv_path))

  # Windowsで使えない文字を置換
  fn <- gsub("[<>:\"/\\\\|?*\\[\\]]", "_", fn)

  file.path(output_dir, paste0(fn, ".xlsx"))
}

# 文字コードを変えながらCSVを読み込む
read_csv_safely <- function(csv_path) {
  encodings <- c("UTF-8", "CP932", "Shift-JIS")

  last_error <- NULL

  for (enc in encodings) {
    result <- try(
      readr::read_csv(
        file = csv_path,
        locale = readr::locale(encoding = enc),
        show_col_types = FALSE,
        guess_max = 10000,
        name_repair = "unique"
      ),
      silent = TRUE
    )

    if (!inherits(result, "try-error")) {
      attr(result, "used_encoding") <- enc
      return(result)
    }

    last_error <- result
  }

  stop("CSVの読み込みに失敗しました: ", basename(csv_path), "\n", as.character(last_error))
}

# ===== 実行：失敗しても次へ =====

ok_list <- character(0)
ng_list <- list()

for (csv in csv_files) {
  xlsx <- make_xlsx_path(csv)

  message("Converting: ", basename(csv), " -> ", basename(xlsx))

  res <- try({
    df <- read_csv_safely(csv)

    used_encoding <- attr(df, "used_encoding")
    message("  Encoding: ", used_encoding)

    # writexl用に通常のdata.frameへ変換
    df <- as.data.frame(df)

    # 既存ファイルがあれば削除
    if (file.exists(xlsx)) {
      unlink(xlsx)
    }

    write_xlsx(
      x = list(Sheet1 = df),
      path = xlsx
    )

    if (!file.exists(xlsx)) {
      stop("変換は実行されましたが、出力ファイルが見つかりません。")
    }

    TRUE
  }, silent = TRUE)

  if (inherits(res, "try-error")) {
    msg <- as.character(res)
    message("  ❌ Failed: ", msg)
    ng_list[[basename(csv)]] <- msg
  } else {
    message("  ✅ Saved: ", xlsx)
    ok_list <- c(ok_list, basename(xlsx))
  }
}

message("\n=== SUMMARY ===")
message(sprintf("成功: %d 件 / 失敗: %d 件", length(ok_list), length(ng_list)))

if (length(ok_list)) {
  message("  OK: ", paste(ok_list, collapse = ", "))
}

if (length(ng_list)) {
  message("  NG: ")
  for (nm in names(ng_list)) {
    message("    - ", nm, " : ", ng_list[[nm]])
  }
}
