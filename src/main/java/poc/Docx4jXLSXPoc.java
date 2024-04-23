package poc;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.JaxbSmlPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorkbookPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.xlsx4j.exceptions.Xlsx4jException;
import org.xlsx4j.org.apache.poi.ss.usermodel.DataFormatter;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.SheetData;
import org.xlsx4j.sml.Worksheet;

import jakarta.xml.bind.JAXBException;

public class Docx4jXLSXPoc {

    private static final DataFormatter formatter = new DataFormatter();

    public static void main(String[] args) throws Exception {

        String inputfilepath = "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\test.xlsx";

        String outputfilepath = "xlsx_test_out.xlsx";

        SpreadsheetMLPackage mlPackage = SpreadsheetMLPackage.load(new File(inputfilepath));


        // Replace variables with text
        replaceVariables(mlPackage);

        mlPackage.save(new File(outputfilepath));

        // get data from sheet
        getClientInfo(mlPackage);
    }

    private static void replaceVariables(final SpreadsheetMLPackage mlPackage) throws Docx4JException, JAXBException {
        // Be sure to get the part which actually contains your variables!
        JaxbSmlPart smlPart = (JaxbSmlPart) mlPackage.getParts().get(new PartName("/xl/sharedStrings.xml"));

        // Replace variables
        HashMap<String, String> mappings = new HashMap<>();
        mappings.put("SIGNER_NAME_1", "Jan Kowalski");
        mappings.put("SIGNER_NAME_2", "Piotr Nowak");
        mappings.put("SIGNER_NAME_3", "Zbigniew Reczek");
        mappings.put("SIGNER_NAME_4", "Krzysztof Futro");
        smlPart.variableReplace(mappings);
    }

    private static void getClientInfo(final SpreadsheetMLPackage mlPackage) throws Xlsx4jException, Docx4JException {
        WorkbookPart workbookPart = mlPackage.getWorkbookPart();
        WorksheetPart sheet = workbookPart.getWorksheet(0);

        Worksheet ws = sheet.getContents();
        SheetData sheetData = ws.getSheetData();

        Map<String, String> clientMap = Map.of(
                "firstName", "B1",
                "lastName", "B2",
                "email", "B3",
                "phone", "B4"
        );

        Client client = Client.builder()
                .firstName(getCell(sheetData, clientMap.get("firstName")))
                .lastName(getCell(sheetData, clientMap.get("lastName")))
                .email(getCell(sheetData, clientMap.get("email")))
                .phone(getCell(sheetData, clientMap.get("phone")))
                .build();

        System.out.println(client);
    }

    private static String getCell(final SheetData sheetData, final String cell) {
        for (Row row : sheetData.getRow()) {
            for (Cell c : row.getC()) {
                if (StringUtils.equals(c.getR(), cell)) {
                    return formatter.formatCellValue(c);
                }
            }
        }
        throw new RuntimeException("Cell content not found for cell: " + cell);
    }

}