{-# LANGUAGE OverloadedStrings #-}

import Codec.Xlsx
import Control.Lens
import Data.Time.Calendar
import Data.Time.Calendar.MonthDay
import Data.Time.Clock
import System.IO
import Text.PDF
import Text.PDF.Info

modifyInvoiceXlsx :: FilePath -> String -> String -> IO Xlsx
modifyInvoiceXlsx inputFileName numberCell dateCell = do
    -- Load the XLSX file
    bs <- B.readFile inputFileName
    let sheets = toXlsx bs
        modifyCell cellName f = over cellValue f $ sheets ^?! ixSheet "Sheet1" . ixCell (cellNameToRef cellName)
        
    -- Get the current invoice number and date
    let invoiceNumber = modifyCell numberCell _cellValueNumber
        invoiceDate = modifyCell dateCell (addGregorianMonthsClip 1 . _cellValueDay)
    
    -- Modify the invoice number and date
    let modifiedSheets = setCellValue (cellNameToRef numberCell) (CellNumber (invoiceNumber + 1)) $
                         setCellValue (cellNameToRef dateCell) (CellDay invoiceDate) sheets
    
    pure $ toXlsx modifiedSheets

-- Convert the XLSX file to a PDF
convertToPdf :: FilePath -> FilePath -> IO ()
convertToPdf inputFileName outputFileName = do
    -- Load the XLSX content
    bs <- B.readFile inputFileName
    let sheets = toXlsx bs
    
    -- Create the PDF
    let pdfContent = execPDF $ do
          setInfoAttribute Author "Your Name"
          setInfoAttribute Title "Invoice"
          setInfoAttribute CreationDate (utctDayTime (secondsToDiffTime 0))

          -- Draw the content on the PDF
          let drawLine y text = do
                drawText $ printBox (80, y) 300 12 (textDefault text)
                pure (y - 20)

          -- Start drawing from (100, 750)
          let (_, y) = foldl (\(x, y) line -> (x, y - 20)) (100, 750) $ map (_cellValueText . _cellValue) (getValuesFromXlsx sheets "Sheet1")
          _ <- foldM drawLine y (map (_cellValueText . _cellValue) (getValuesFromXlsx sheets "Sheet1"))

    -- Save the PDF
    withBinaryFile outputFileName WriteMode $ \h ->
        writePDF h pdfContent

main :: IO ()
main = do
    let inputFileName = "invoice_template.xlsx"
        outputFileName = "modified_invoice.pdf"

    -- Step 1: Modify the XLSX file
    modifiedInvoice <- modifyInvoiceXlsx inputFileName

    -- Save the modified XLSX file (optional)
    let modifiedFileName = "modified_invoice.xlsx"
    B.writeFile modifiedFileName $ fromXlsx modifiedInvoice

    -- Step 2: Convert the modified XLSX to PDF
    convertToPdf modifiedFileName outputFileName

    -- Optionally, remove the temporary modified XLSX file
    removeFile modifiedFileName

    putStrLn "Invoice modified and saved as PDF successfully!"
