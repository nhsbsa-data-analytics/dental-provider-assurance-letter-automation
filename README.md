# Dental Provider Assurance - Letter Automation

This process can be automated through a mixture of Microsoft Word Mail Merge, Macros and R code. The steps are as follows:

1. Create the input file
2. Produce the individual `.pdf` files
3. Run the R script

## Create the input file

This should be a `.csv` file with one row per contractor and column names formatted in `snake_case` to make the mail merge document easier to develop.

## Produce the individual `.pdf` files

A mail merge document must be created and linked to the input file in order to produce the `.pdf` files. 

A `template.docx` file has been included with some common design patterns used when creating dynamic mail merges.

You can link this to the input file by adding the two macros:
1. `RunMailMerge.bas` - output an individual `.pdf` for each row of the input file.
2. `DeleteEmptyRowsSkipIfMerged.bas` - Delete empty rows from a table in the mail merge document (remove line 45 in `RunMailMerge.bas` if not needed).

and defining the parameters on lines 7-9 in `RunMailMerge.bas`. This should help preview the file as you are going and ensure that the merge fields are correct.

When you are happy run `RunMailMerge` from inside Microsoft Word to produce the individual `.pdf` files.

Note: You can tweak line 30 in `RunMailMerge.bas` to a sensible number during development (e.g. `For recordNumber = 1 To 10` will produce 10 outputs).

## Merge the `.pdf` files by STP 

A simple R script `merge_pdfs.R` is provided to merge the individual `.pdf` files into larger ones grouped by STP.

As before, there are parameters to define at the start of the script (lines 4-6).
