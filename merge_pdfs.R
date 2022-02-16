library(dplyr)

# Read input data as a dataframe
df <- read.csv(file = "FINAL LETTER PRODUCTION CAT 1 AND 2 ONLY v3.csv")

# Combine PDF files by STP code
df %>%
  # Get the path of the PDF files
  mutate(
    pdf_file_path = paste0(
      "output/BNPR_D20201104_YP192012",
      "_B", latest_provider_number,
      "_C", gsub("/", "", contract_number),
      ".pdf"
    )
  ) %>%
  # Group by STP code and combine the PDF files
  group_by(stp_code) %>%
  group_walk(
    .f = ~ qpdf::pdf_combine(
      input = .x$pdf_file_path,
      output = paste0("output/BNPR_D20201104_YP192012_H", .y, ".pdf")
    )
  )
