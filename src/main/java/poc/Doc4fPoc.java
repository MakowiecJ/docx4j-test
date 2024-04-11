package poc;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Document;
import org.docx4j.wml.P;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Doc4fPoc {

    private static final boolean DELETE_BOOKMARK = false;
    protected static Logger log = LoggerFactory.getLogger(Doc4fPoc.class);
    private static final org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();


    public static void main(String[] args) throws Exception {

        // Load the .docx file
//        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File("C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\TEST.docx"));
        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File("C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\brd_test_original.docx"));
//        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File("C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\bookmarks_replaced.docx"));
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

        // Replace bookmarks
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
        wordMLPackage.save(new File("bookmarks_replaced.docx"));

//         Find all bookmarks
//        List<Object> content = mainDocumentPart.getContent();
//        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
//        RangeFinder rt = new RangeFinder();
//        new TraversalUtil(content, rt);
//
//        for (CTBookmark bm : rt.getStarts()) {
//            System.out.println("Bookmark Name: " + bm.getName());
//        }

        // Export to pdf

        OutputStream pdfOutput = new FileOutputStream("output.pdf");
        FOSettings foSettings = new FOSettings(wordMLPackage);
        FopFactoryBuilder fopFactoryBuilder = FORendererApacheFOP.getFopFactoryBuilder(foSettings);
        FopFactory fopFactory = fopFactoryBuilder.build();
//        foSettings.setFopConfig(fopFactory.newFop(MimeConstants.MIME_PDF, pdfOutput));

        FOUserAgent foUserAgent = FORendererApacheFOP.getFOUserAgent(foSettings, fopFactory);
        // configure foUserAgent as desired
        foUserAgent.setTitle("my title");
        foUserAgent.getRendererOptions().put("version", "2.0");

        Docx4J.toFO(foSettings, pdfOutput, Docx4J.FLAG_EXPORT_PREFER_XSL);

        pdfOutput.flush();
        pdfOutput.close();
    }

    private void replaceBookmarkContents(List<Object> paragraphs, Map<DataFieldName, String> data) throws Exception {

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

}
