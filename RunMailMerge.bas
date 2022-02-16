' Adapted from https://www.msofficeforums.com/word-vba/46987-delete-blank-table-rows-merged-document-split.html


Option Explicit

' Define parameters here
Const INPUT_FILE As String = "path\to\input-file.csv" ' Must be a .csv as word has a 255 character limit on fields from .xlsx
Const OUTPUT_DIR As String = "path\to\output-dir\" ' Include \ at end and ensure directory exists
Const FILENAME As String = "" ' e.g. BNPR_D20201104_YP192012 or AYRPR_D20211014_YP2122


Sub RunMailMerge()

    Dim MainDoc As Document, TargetDoc As Document
    Dim recordNumber As Long, totalRecord As Long
    Dim providerNumber As String, contractNumber As String
        
    Set MainDoc = ActiveDocument
    With MainDoc.MailMerge
            
            ' Read in the input file as a datasource
            .OpenDataSource Name:=INPUT_FILE
            
            ' Find the number of rows in the input file (.RecordCount won't work)
            With CreateObject("Scripting.FileSystemObject")
                totalRecord = UBound(Split(.OpenTextFile(INPUT_FILE, 1).ReadAll, vbNewLine)) - 1
            End With
            
            ' Loop over each record
            For recordNumber = 1 To 3
                
                ' One record at a time
                With .DataSource
                    .ActiveRecord = recordNumber
                    .FirstRecord = recordNumber
                    .LastRecord = recordNumber
                End With
                
                .Destination = wdSendToNewDocument
                .Execute False
                
                Set TargetDoc = ActiveDocument
                
                ' Call this module if there are tables with empty rows that need removed (e.g. the final table)
                Call DeleteEmptyRowsSkipIfMerged(TargetDoc)
                
                ' Pull the parameters used to name the files
                providerNumber = .DataSource.DataFields("latest_provider_number").Value
                contractNumber = Replace(.DataSource.DataFields("contract_number").Value, "/", "")
                
                ' Produce the output file
                TargetDoc.ExportAsFixedFormat OutputFileName:=OUTPUT_DIR & FILENAME & "_B" & providerNumber & "_C" & contractNumber & ".pdf", ExportFormat:=wdExportFormatPDF
                TargetDoc.Close False
                
                Set TargetDoc = Nothing
                        
            Next recordNumber
    
    End With
    
    Set MainDoc = Nothing
    
End Sub
