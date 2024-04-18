package poc;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.FopFactoryBuilder;
import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.fo.renderers.FORendererApacheFOP;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
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
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGrid;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import jakarta.xml.bind.JAXBElement;

public class Doc4fPoc {

    private static final boolean DELETE_BOOKMARK = false;
    protected static Logger log = LoggerFactory.getLogger(Doc4fPoc.class);
    private static final org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();


    public static void main(String[] args) throws Exception {


        // Load the .docx file
        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File("C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\test_document_with_bookmarks.docx"));
        WordprocessingMLPackage secondDocPackage = Docx4J.load(new File("C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\TEST.docx"));
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        // adding sub document link
        JAXBElement<CTRel> subdoc = createSubdocLink(mainDocumentPart, "C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\TEST.docx");
        org.docx4j.wml.ObjectFactory wmlFactory = Context.getWmlObjectFactory();
        org.docx4j.wml.P paragraph = wmlFactory.createP();
        paragraph.getContent().add( subdoc );
        mainDocumentPart.addObject(paragraph);
        // coping subdocument content
        copySubDocumentContent(wordMLPackage, secondDocPackage);

        // Find all bookmarks
        findAllBookmarks(mainDocumentPart);

        // Replace bookmarks with text and tables
        replaceBookmarks(wordMLPackage, mainDocumentPart);

        // Save the modified Word document
        wordMLPackage.save(new File("output.docx"));

        long startTime = System.currentTimeMillis();
        // Export to pdf
        exportToPdf(wordMLPackage);

        long endTime = System.currentTimeMillis();
        long duration = (endTime - startTime);
        log.info("Generating PDF duration: " + duration + "ms");
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
            rel.setType( Namespaces.SUBDOCUMENT  );
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

    private static void exportToPdf(final WordprocessingMLPackage wordMLPackage) throws Exception {
        OutputStream pdfOutput = new FileOutputStream("output.pdf");

        // use FO converter
        FOSettings foSettings = new FOSettings(wordMLPackage);
        FopFactoryBuilder fopFactoryBuilder = FORendererApacheFOP.getFopFactoryBuilder(foSettings);
        FopFactory fopFactory = fopFactoryBuilder.build();

        FOUserAgent foUserAgent = FORendererApacheFOP.getFOUserAgent(foSettings, fopFactory);
        // configure foUserAgent
        foUserAgent.setTitle("my title");
        foUserAgent.getRendererOptions().put("version", "2.0");

        Docx4J.toFO(foSettings, pdfOutput, Docx4J.FLAG_EXPORT_PREFER_XSL);
//        Docx4J.toFO(foSettings, pdfOutput, Docx4J.FLAG_EXPORT_PREFER_NONXSL); // little bit less demanding option

        // Use word/powerpoint to export to pdf
//        Documents4jLocalServices exporter = new Documents4jLocalServices();
//        exporter.export(wordMLPackage, pdfOutput);

        pdfOutput.flush();
        pdfOutput.close();
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

        Doc4fPoc bti = new Doc4fPoc();

        bti.replaceBookmarkContents(body.getContent(), map);

        // Replacing bookmark with some table
        bti.replaceBookmarkWithTable(body, "TEST_LIST");
    }

    private void replaceBookmarkContents(final List<Object> paragraphs, final Map<DataFieldName, String> data) throws Exception {

        RangeFinder rt = new RangeFinder();
//        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
        new TraversalUtil(paragraphs, rt);

        for (CTBookmark bm : rt.getStarts()) {

            // do we have data for this one?
            if (bm.getName() == null) continue;
            String value = data.get(new DataFieldName(bm.getName()));
            if (value == null) continue;

            try {
                // Can't just remove the object from the parent,
                // since in the parent, it may be wrapped in a JAXBElement
                List<Object> theList = null;
                if (bm.getParent() instanceof P) {
                    theList = ((ContentAccessor) (bm.getParent())).getContent();
                } else {
                    continue;
                }

                int rangeStart = -1;
                int rangeEnd = -1;
                int i = 0;
                for (Object ox : theList) {
                    Object listEntry = XmlUtils.unwrap(ox);
                    if (listEntry.equals(bm)) {
                        if (DELETE_BOOKMARK) {
                            rangeStart = i;
                        } else {
                            rangeStart = i + 1;
                        }
                    } else if (listEntry instanceof CTMarkupRange) {
                        if (((CTMarkupRange) listEntry).getId().equals(bm.getId())) {
                            if (DELETE_BOOKMARK) {
                                rangeEnd = i;
                            } else {
                                rangeEnd = i > rangeStart ? i - 1 : i;     // handle empty bookmark case
                            }
                            break;
                        }
                    }
                    i++;
                }

                if (rangeStart > 0 && rangeEnd >= rangeStart) {

                    // Delete the bookmark range
                    if (rangeEnd > rangeStart) {
                        for (int j = rangeEnd; j >= rangeStart; j--) {
                            theList.remove(j);
                        }
                    }

                    // Delete field before bookmark
                    theList.remove(0);

                    // now add a run
                    org.docx4j.wml.R run = factory.createR();
                    org.docx4j.wml.Text t = factory.createText();
                    run.getContent().add(t);
                    t.setValue(value);

                    theList.add(rangeStart, run);
                }

            } catch (ClassCastException cce) {
                log.error(cce.getMessage(), cce);
            }
        }
    }

    private void replaceBookmarkWithTable(final Body body, final String bookmarkName) throws Exception {
        var paragraphs = body.getContent();
        RangeFinder rt = new RangeFinder();
        new TraversalUtil(paragraphs, rt);

        for (CTBookmark bm : rt.getStarts()) {
            // Check if the bookmark has a name and if there's bookmarkName for it
            if (bm.getName() == null || !Objects.equals(bm.getName(), bookmarkName)) {
                continue;
            }
            try {
                // Get the parent of the paragraph


                // Create a table to replace the paragraph
                Tbl table = createClientTable(List.of(
                        List.of("Name", "Phone Number", "Email"),
                        List.of("Alice Johnson", "555-123-4567", "alice@example.com"),
                        List.of("Bob Smith", "555-987-6543", "bob@example.com")
                        // Add more client data as needed
                ));

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

    public static Tbl createClientTable(List<List<String>> clients) {
        try {
            // Create a table
            Tbl table = TblFactory.createTable(clients.size(), clients.get(0).size(), 4000);

            // setting table grid and style taken from http://webapp.docx4java.org/OnlineDemo/PartsList.html generated
            TblPr tblpr = factory.createTblPr();
            // Create object for tblStyle
            CTTblPrBase.TblStyle tblprbasetblstyle = factory.createCTTblPrBaseTblStyle();
            tblpr.setTblStyle(tblprbasetblstyle);
            tblprbasetblstyle.setVal( "Tabela-Siatka");
            // Create object for tblW
            TblWidth tblwidth = factory.createTblWidth();
            tblpr.setTblW(tblwidth);
            tblwidth.setW( BigInteger.valueOf( 0) );
            tblwidth.setType( "auto");
            table.setTblPr(tblpr);


            table.setTblGrid(new TblGrid());
            List<Object> rows = table.getContent();

            for (int i = 0; i < clients.size(); i++) {
                var client = clients.get(i);
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
