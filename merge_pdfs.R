library(dplyr)

# Define parameters here
INPUT_FILE = "path/to/input-file.csv" # Must be a .csv as word has a 255 character limit on fields from .xlsx
OUTPUT_DIR = "path/to/output-dir/" # Include / at end and ensure directory exists
FILENAME = "" # e.g. BNPR_D20201104_YP192012 or AYRPR_D20211014_YP2122

# Read input data as a dataframe
df <- read.csv(file = INPUT_FILE)

# Combine PDF files by STP code
df %>%
  # Get the path of the PDF files
  mutate(
    pdf_file_path = paste0(
      OUTPUT_DIR,
      FILENAME,
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
      output = paste0(OUTPUT_DIR, FILENAME, "_H", .y, ".pdf")
    )
  )
