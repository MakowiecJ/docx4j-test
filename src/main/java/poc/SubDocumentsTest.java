package poc;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Document;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import poc.converters.LibreConverter;

public class SubDocumentsTest {

    protected static Logger log = LoggerFactory.getLogger(Docx4jPoc.class);
    private static final org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();


    public static void main(String[] args) throws Exception {
        String inputFilePath = "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\test_document.docx";
//        Converter pdfConverter = new FOConverter();
        LibreConverter pdfConverter = new LibreConverter();
        WordprocessingMLPackage wordMLPackage1 = Docx4J.load(new File(inputFilePath));
        WordprocessingMLPackage wordMLPackage2 = Docx4J.load(new File(inputFilePath));
        MainDocumentPart mainDocumentPart1 = wordMLPackage1.getMainDocumentPart();
        MainDocumentPart mainDocumentPart2 = wordMLPackage2.getMainDocumentPart();


        int sampleSize = 100;
        long startTime = System.currentTimeMillis();
        for (int i = 0; i < sampleSize; i++) {
            replaceBookmarks(mainDocumentPart1);
            copySubDocumentContent(wordMLPackage1, wordMLPackage2);
            replaceBookmarks(mainDocumentPart1);
        }
        long endTime = System.currentTimeMillis();
        long duration = endTime - startTime;


        // Save the modified Word document
        wordMLPackage1.save(new File("subDocumentsTestOutput.docx"));
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        wordMLPackage1.save(byteArrayOutputStream);
        ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());

        OutputStream pdfOutputStream = new FileOutputStream("subDocumentsTestOutput.pdf");

        long pdfStartTime = System.currentTimeMillis();
        pdfConverter.convert(byteArrayInputStream, pdfOutputStream);
        long pdfEndTime = System.currentTimeMillis();
        long pdfDuration = pdfEndTime - pdfStartTime;

        // Needed to stop libre office!
        pdfConverter.stopOffice();

        System.out.println("Merging " + sampleSize + " templates: " + duration + "ms");
        System.out.println("Generating PDF duration: " + pdfDuration + "ms");
    }

    private static void copySubDocumentContent(final WordprocessingMLPackage firstDocPackage, final WordprocessingMLPackage secondDocPackage) throws Exception {
        // Get the body of the second document
        List<Object> secondBody = secondDocPackage.getMainDocumentPart().getJAXBNodesViaXPath("//w:body", false);

        // Append the second document's body to the first document
        for (Object body : secondBody) {
            List<Object> children = ((org.docx4j.wml.Body) body).getContent();
            for (Object child : children) {
                // Add each element to the first document
                firstDocPackage.getMainDocumentPart().getContent().add(child);
            }
        }
    }

    private static void replaceBookmarks(final MainDocumentPart mainDocumentPart) throws Exception {
        Map<DataFieldName, String> map = new HashMap<>();
        map.put(new DataFieldName("AUTH_1"), "Exposure 1");
        map.put(new DataFieldName("AUTH_1_5"), "Exposure 2");
        map.put(new DataFieldName("AUTH_5_7"), "Exposure 3");
        map.put(new DataFieldName("AUTH_7_10"), "Exposure 4");
        map.put(new DataFieldName("AUTH_10_15"), "Exposure 5");
        map.put(new DataFieldName("AUTH_15"), "Exposure 6");
        map.put(new DataFieldName("CLN_RDO"), "Client RDO");
        map.put(new DataFieldName("CLN_CLASSICAL"), "Client classical");
        map.put(new DataFieldName("CLN_SG_L"), "Client total SG lease");
        map.put(new DataFieldName("CLN_FACT"), "Client total fact");
        map.put(new DataFieldName("CLN_INDIVIDUAL"), "Client related individuals");
        map.put(new DataFieldName("CLN_RLI"), "Client total RLI");
        map.put(new DataFieldName("CLN_CVAR"), "Client total CVaR");
        map.put(new DataFieldName("ECG_RDO"), "ECG RDO");
        map.put(new DataFieldName("ECG_CLASSICAL"), "ECG classical");
        map.put(new DataFieldName("ECG_SG_L"), "ECG total SG lease");
        map.put(new DataFieldName("ECG_FACT"), "ECG total fact");
        map.put(new DataFieldName("ECG_INDIVIDUAL"), "ECG related individuals");
        map.put(new DataFieldName("ECG_RLI"), "ECG total RLI");
        map.put(new DataFieldName("ECG_CVAR"), "ECG total CVaR");
        map.put(new DataFieldName("SIGNER_NAME_1"), "Jan Kowalski");
        map.put(new DataFieldName("SIGNER_NAME_2"), "Piotr Nowak");
        map.put(new DataFieldName("SIGNER_NAME_3"), "Zbigniew Reczek");
        map.put(new DataFieldName("SIGNER_NAME_4"), "Krzysztof Futro");
        map.put(new DataFieldName("SIGNER_POSITION_1"), "CEO");
        map.put(new DataFieldName("SIGNER_POSITION_2"), "Manager");
        map.put(new DataFieldName("SIGNER_POSITION_3"), "Product Owner");
        map.put(new DataFieldName("SIGNER_POSITION_4"), "Product Owner");
        map.put(new DataFieldName("APPROVAL_NAME"), "Jan Kowalski");
        map.put(new DataFieldName("APPROVAL_POSITION"), "CEO");

        Document wmlDocumentEl = mainDocumentPart.getJaxbElement();
        Body body = wmlDocumentEl.getBody();

        replaceBookmarkContents(body.getContent(), map);
    }

    private static void replaceBookmarkContents(final List<Object> paragraphs, final Map<DataFieldName, String> data) throws Exception {

        RangeFinder rt = new RangeFinder();
        new TraversalUtil(paragraphs, rt);

        for (CTBookmark bm : rt.getStarts()) {

            if (bm.getName() == null) continue;
            String value = data.get(new DataFieldName(bm.getName()));
            if (value == null) continue;

            try {
                List<Object> theList = null;
                if (bm.getParent() instanceof P) {
                    theList = ((ContentAccessor) (bm.getParent())).getContent();
                } else {
                    continue;
                }

                // Copy run formatting
                RPr rpr = ((R) theList.get(0)).getRPr();


                theList.clear();


                // now add a run
                org.docx4j.wml.R run = factory.createR();
                org.docx4j.wml.Text t = factory.createText();
                run.getContent().add(rpr);
                run.getContent().add(t);
                t.setValue(value);

                theList.add(run);


            } catch (ClassCastException cce) {
                log.error(cce.getMessage(), cce);
            }
        }
    }


}
