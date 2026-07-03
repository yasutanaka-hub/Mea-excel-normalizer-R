# CSVファイルをExcel形式（.xlsx）に一括変換するMac/Windows共通版
# 区切り文字を自動判定し、1行目もデータとして保持する版

# ===== 必要パッケージ =====
required_packages <- c("writexl")

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
}

library(writexl)

# ===== 設定 =====

# Mac / Windows 共通：ユーザーフォルダ/Desktop/R を使う
base_dir <- file.path(path.expand("~"), "Desktop", "R")

input_dir  <- file.path(base_dir, "csv_files")
output_dir <- file.path(base_dir, "excel_conv_output_files")

if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

if (!dir.exists(input_dir)) {
  stop("入力フォルダが見つかりません: ", input_dir)
}

csv_files <- list.files(
  input_dir,
  pattern = "\\.[Cc][Ss][Vv]$",
  full.names = TRUE
)

csv_files <- csv_files[!grepl("(^~\\$)|(^\\._)", basename(csv_files))]

if (length(csv_files) == 0) {
  stop(sprintf("CSVが見つかりませんでした: %s", input_dir))
}

# ===== 共通ユーティリティ =====

make_xlsx_path <- function(csv_path) {
  fn <- tools::file_path_sans_ext(basename(csv_path))

  # Windows / macOS で問題になりやすい文字を置換
  fn <- gsub("[<>:\"/\\\\|?*\\[\\]]", "_", fn)

  file.path(output_dir, paste0(fn, ".xlsx"))
}

# 区切り文字を推定する
detect_delimiter <- function(csv_path, encoding) {
  candidates <- c("," = ",", "tab" = "\t", ";" = ";", "|" = "|")

  scores <- sapply(candidates, function(sep) {
    con <- file(csv_path, open = "r", encoding = encoding)
    on.exit(close(con), add = TRUE)

    fields <- tryCatch(
      count.fields(
        con,
        sep = sep,
        quote = "\"",
        blank.lines.skip = FALSE,
        comment.char = ""
      ),
      error = function(e) integer(0)
    )

    fields <- fields[!is.na(fields)]

    if (length(fields) == 0) {
      return(0)
    }

    median(fields)
  })

  best <- names(scores)[which.max(scores)]

  if (is.na(best) || scores[[best]] <= 1) {
    return(",")
  }

  candidates[[best]]
}

# CSVをセル配置優先で読む
read_csv_keep_cells <- function(csv_path) {
  encodings <- c("UTF-8-BOM", "UTF-8", "CP932", "Shift-JIS")

  last_error <- NULL

  for (enc in encodings) {
    result <- tryCatch({
      sep <- detect_delimiter(csv_path, enc)

      con <- file(csv_path, open = "r", encoding = enc)
      on.exit(close(con), add = TRUE)

      field_counts <- count.fields(
        con,
        sep = sep,
        quote = "\"",
        blank.lines.skip = FALSE,
        comment.char = ""
      )

      field_counts <- field_counts[!is.na(field_counts)]

      if (length(field_counts) == 0) {
        stop("列数を判定できませんでした。")
      }

      max_cols <- max(field_counts)

      df <- read.table(
        file = csv_path,
        sep = sep,
        quote = "\"",
        header = FALSE,
        fill = TRUE,
        col.names = paste0("X", seq_len(max_cols)),
        colClasses = "character",
        stringsAsFactors = FALSE,
        check.names = FALSE,
        comment.char = "",
        blank.lines.skip = FALSE,
        na.strings = "",
        fileEncoding = enc
      )

      attr(df, "used_encoding") <- enc
      attr(df, "used_separator") <- ifelse(sep == "\t", "TAB", sep)

      df
    }, error = function(e) {
      last_error <<- e
      NULL
    })

    if (!is.null(result)) {
      return(result)
    }
  }

  stop("CSVの読み込みに失敗しました: ", basename(csv_path), "\n", last_error$message)
}

# ===== 実行 =====

ok_list <- character(0)
ng_list <- list()

for (csv in csv_files) {
  xlsx <- make_xlsx_path(csv)

  message("Converting: ", basename(csv), " -> ", basename(xlsx))

  res <- try({
    df <- read_csv_keep_cells(csv)

    used_encoding <- attr(df, "used_encoding")
    used_separator <- attr(df, "used_separator")

    message("  Encoding: ", used_encoding)
    message("  Separator: ", used_separator)

    if (file.exists(xlsx)) {
      unlink(xlsx)
    }

    write_xlsx(
      x = list(Sheet1 = df),
      path = xlsx,
      col_names = FALSE
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
