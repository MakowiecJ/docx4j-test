package poc;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.CTRel;
import org.docx4j.wml.CTTblPrBase;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Document;
import org.docx4j.wml.Jc;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import jakarta.xml.bind.JAXBElement;
import jakarta.xml.bind.JAXBException;
import poc.converters.Converter;
import poc.converters.FOConverter;
import poc.converters.LibreConverter;

public class Docx4jPoc {

    private static final boolean DELETE_BOOKMARK = false;
    protected static Logger log = LoggerFactory.getLogger(Docx4jPoc.class);
    private static final org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();


    public static void main(String[] args) throws Exception {

        LibreConverter pdfConverter = new LibreConverter();

        String inputFilePath = "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\test_document.docx";
        String inputFilePath2 = "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\TEST.docx";

        // Load the .docx file
        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(inputFilePath));
        WordprocessingMLPackage secondDocPackage = Docx4J.load(new File(inputFilePath2));
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        // adding sub document link (can break pdf generation!)
//        JAXBElement<CTRel> subdoc = createSubdocLink(mainDocumentPart, "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\TEST.docx");
//        org.docx4j.wml.ObjectFactory wmlFactory = Context.getWmlObjectFactory();
//        org.docx4j.wml.P paragraph = wmlFactory.createP();
//        paragraph.getContent().add(subdoc);
//        mainDocumentPart.addObject(paragraph);

        // copying subdocument content
        copySubDocumentContent(wordMLPackage, secondDocPackage);

        // Find all bookmarks
        findAllBookmarks(mainDocumentPart);

        // Replace bookmarks with text
        replaceBookmarks(wordMLPackage, mainDocumentPart);

        // Create table out of list - each list represents a table row
        Tbl table1 = createTable(List.of(
                List.of("Name", "Phone Number", "Email", "Test"),
                List.of("Alice Johnson", "555-123-4567", "alice@example.com", "test1"),
                List.of("Bob Smith", "555-987-6543", "bob@example.com", "test2")
        ));

        // Replace bookmarks with table
        replaceBookmarkWithTable(mainDocumentPart.getContents().getBody(), "BOOKMARK_TABLE", table1);

        // Replace variables with text
        HashMap<String, String> mappings = new HashMap<>();
        mappings.put("NAME_1", "Jan Kowalski");
        mappings.put("NAME_2", "Jan Kowalski");
        mappings.put("POSITION_2", "CEO");
        mappings.put("POSITION_2", "PO");
        mainDocumentPart.variableReplace(mappings);

        // Replace variable with table
        Tbl table2 = createTable(List.of(
                List.of("Name", "Phone Number"),
                List.of("Alice Johnson", "555-123-4567"),
                List.of("Bob Smith", "555-987-6543"),
                List.of("Anna Stone", "414-251-1234")
        ));
        replaceVariableWithTable(mainDocumentPart, "VARIABLE_TABLE", table2);


        // Save the modified Word document
        wordMLPackage.save(new File("output.docx"));
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        wordMLPackage.save(byteArrayOutputStream);
        ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());

        OutputStream pdfOutputStream = new FileOutputStream("output.pdf");

        pdfConverter.convert(byteArrayInputStream, pdfOutputStream);

        // Needed for libre office converter to stop!
        pdfConverter.stopOffice();

    }

    public static void replaceVariableWithTable(final MainDocumentPart mainDocumentPart, final String variableName, final Tbl table)
            throws JAXBException, XPathBinderAssociationIsPartialException {

        String xpath = "//w:t";
        List<Object> list = mainDocumentPart.getJAXBNodesViaXPath(xpath, false);
        for (Object o : list) {
            if (o instanceof JAXBElement) {
                JAXBElement element = (JAXBElement) o;
                Text text = (Text) element.getValue();
                if (!text.getValue().contains(variableName)) {
                    continue;
                }
                R r = (R) text.getParent();
                P p = (P) r.getParent();
                Body b = (Body) p.getParent();
                int index = b.getContent().indexOf(p);
                if (index != -1 && index < b.getContent().size() - 1) {
                    // Replace the paragraph with the table
                    b.getContent().remove(index); // Remove the paragraph
                    b.getContent().add(index, table); // Insert the table in place of the paragraph
                }
            }
        }
    }

    public static JAXBElement<CTRel> createSubdocLink(MainDocumentPart mdp,
                                                      String subdocName) {

        try {

            // We need to add a relationship to word/_rels/document.xml.rels
            // but since its external, we don't use the
            // usual wordMLPackage.getMainDocumentPart().addTargetPart
            // mechanism
            org.docx4j.relationships.ObjectFactory factory =
                    new org.docx4j.relationships.ObjectFactory();

            org.docx4j.relationships.Relationship rel = factory.createRelationship();
            rel.setType(Namespaces.SUBDOCUMENT);
            rel.setTarget(subdocName);
            rel.setTargetMode("External");

            mdp.getRelationshipsPart().addRelationship(rel);

            // addRelationship sets the rel's @Id


            org.docx4j.wml.ObjectFactory wmlOF = new org.docx4j.wml.ObjectFactory();

            CTRel ctRel = wmlOF.createCTRel();
            ctRel.setId(rel.getId());

            return wmlOF.createPSubDoc(ctRel);

        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
            return null;
        }


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

    private static void findAllBookmarks(final MainDocumentPart mainDocumentPart) throws Exception {
        List<Object> content = mainDocumentPart.getContent();
        RangeFinder rt = new RangeFinder();
        new TraversalUtil(content, rt);

        for (CTBookmark bm : rt.getStarts()) {
            log.info("Bookmark Name: " + bm.getName());
        }
    }

    private static void replaceBookmarks(final WordprocessingMLPackage wordMLPackage, final MainDocumentPart mainDocumentPart) throws Exception {
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

    private static void replaceBookmarkWithTable(final Body body, final String bookmarkName, final Tbl table) throws Exception {
        var paragraphs = body.getContent();
        RangeFinder rt = new RangeFinder();
        new TraversalUtil(paragraphs, rt);

        for (CTBookmark bm : rt.getStarts()) {
            // Check if the bookmark has a name and if there's bookmarkName for it
            if (bm.getName() == null || !Objects.equals(bm.getName(), bookmarkName)) {
                continue;
            }
            try {
                // Find the index of the paragraph containing the bookmark within the body
                int index = body.getContent().indexOf(bm.getParent());
                if (index != -1 && index < body.getContent().size() - 1) {
                    // Replace the paragraph with the table
                    body.getContent().remove(index); // Remove the paragraph
                    body.getContent().add(index, table); // Insert the table in place of the paragraph
                }

            } catch (ClassCastException cce) {
                log.error(cce.getMessage(), cce);
            }
        }
    }

    public static Tbl createTable(List<List<String>> tableData) {
        try {
            // Create a table
            Tbl table = TblFactory.createTable(tableData.size(), tableData.get(0).size(), 2000);

            // setting table grid and style taken from http://webapp.docx4java.org/OnlineDemo/PartsList.html generated
            TblPr tblpr = factory.createTblPr();
            // Create object for tblStyle
            CTTblPrBase.TblStyle tblprbasetblstyle = factory.createCTTblPrBaseTblStyle();
            tblpr.setTblStyle(tblprbasetblstyle);
            tblprbasetblstyle.setVal("Tabela-Siatka");
            // Create object for tblW
            TblWidth tblwidth = factory.createTblWidth();
            tblpr.setTblW(tblwidth);
            tblwidth.setW(BigInteger.valueOf(0));
            tblwidth.setType("auto");
            table.setTblPr(tblpr);


            List<Object> rows = table.getContent();

            for (int i = 0; i < tableData.size(); i++) {
                var client = tableData.get(i);
                Tr tr = (Tr) rows.get(i);
                List<Object> cells = tr.getContent();
                for (int j = 0; j < client.size(); j++) {
                    Tc td = (Tc) cells.get(j);

                    P p = new P();
                    R r = new R();
                    Text text = new Text();
                    text.setValue(client.get(j));
                    r.getContent().add(text);
                    p.getContent().add(r);

                    // Centering cells, generated using demo webapp
                    PPr ppr = factory.createPPr();
                    Jc jc = factory.createJc();
                    jc.setVal(org.docx4j.wml.JcEnumeration.CENTER);
                    ppr.setJc(jc);
                    p.setPPr(ppr);

                    td.getContent().add(p);
                }
            }

            return table;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

}
