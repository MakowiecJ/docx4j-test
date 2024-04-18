package poc;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorkbookPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xlsx4j.org.apache.poi.ss.usermodel.DataFormatter;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.SheetData;
import org.xlsx4j.sml.Worksheet;

public class CellContentExtractor {

    private static Logger log = LoggerFactory.getLogger(CellContentExtractor.class);

    public static void main(String[] args) throws Exception {

        String inputfilepath = "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\test.xlsx";

        // Open a document from the file system
        SpreadsheetMLPackage xlsxPkg = SpreadsheetMLPackage.load(new java.io.File(inputfilepath));

        WorkbookPart workbookPart = xlsxPkg.getWorkbookPart();
        WorksheetPart sheet = workbookPart.getWorksheet(0);

        DataFormatter formatter = new DataFormatter();

        // Print some cells content
        displayContent(sheet, formatter);
    }


    private static void displayContent(WorksheetPart sheet, DataFormatter formatter) throws Docx4JException {

        Worksheet ws = sheet.getContents();
        SheetData data = ws.getSheetData();

        var cell21 = data.getRow().get(2).getC().get(1);
        System.out.println("Cell 2:1 (B3): " + formatter.formatCellValue(cell21) + "\n");

        for (Row r : data.getRow() ) {
            System.out.println("row " + r.getR() );

            for (Cell c : r.getC() ) {

//	            CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
//	            System.out.print(cellRef.formatAsString());
//	            System.out.print(" - ");

                // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
                String text = formatter.formatCellValue(c);
                System.out.println(c.getR() + " contains " + text);

            }
        }

    }



}