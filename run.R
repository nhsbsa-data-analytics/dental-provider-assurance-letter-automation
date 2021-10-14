# Read input data as a dataframe
df <- readxl::read_excel(
  path = file.path(
    "AYRPR", 
    "input", 
    "FINAL LETTER PRODUCTION CAT 1 AND 2 ONLY v3.xlsx"
  )
)

# Loop over every row and produce the letter
for (i in 1:nrow(head(df, 2))) {
  
  # Filter the df
  row_ <- df[i, ]
  
  # Pull out attributes needed to name the file
  p_no <- row_["Latest Provider Number"]
  c_no <- gsub("/", "", row_["Contract Number"]) # Remove "/"

  # Produce the pdf
  rmarkdown::render(
    input = file.path("AYRPR", "input", "2021 NDCM YE Letter.Rmd"),
    output_file = paste0("AYRPR_D20211014_YP2122_B", p_no, "_C", c_no),
    output_dir = file.path("AYRPR", "output"),
    params = list(
      header = file.path("AYRPR", "input", "header.png"),
      row_ = row_
    )
  )
}
