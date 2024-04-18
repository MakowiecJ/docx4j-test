package poc;

import java.io.File;
import java.util.HashMap;

import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.JaxbSmlPart;

public class XlsVariableReplace {

    public static void main(String[] args) throws Exception {

        // Input xslx has variables in it: ${title1}
        String inputfilepath = "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\test.xlsx";

        boolean save = true;
        String outputfilepath = "OUT_VariableReplace.xlsx";

        SpreadsheetMLPackage opcPackagepkg = SpreadsheetMLPackage.load(new File(inputfilepath));

        // Be sure to get the part which actually contains your variables!
        JaxbSmlPart smlPart = (JaxbSmlPart) opcPackagepkg.getParts().get(new PartName("/xl/sharedStrings.xml"));
        //JaxbSmlPart smlPart = (JaxbSmlPart)opcPackagepkg.getParts().get(new PartName("/xl/worksheets/sheet1.xml"));


        System.out.println("\n\nBEFORE\n\n:" + smlPart.getXML());

        HashMap<String, String> mappings = new HashMap<String, String>();

        mappings.put("someVar", "Replaced");

        smlPart.variableReplace(mappings);


        // Save it
        if (save) {
            opcPackagepkg.save(new File(outputfilepath));
        } else {
            System.out.println("\n\nAFTER\n\n:" + smlPart.getXML());
        }
    }

}