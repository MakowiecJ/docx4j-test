package poc.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import com.spire.doc.FileFormat;
import com.spire.doc.PdfConformanceLevel;
import com.spire.doc.ToPdfParameterList;

public class SpireIceConverter implements Converter {

    @Override
    public void convert(final String inputFilePath, final String outputFilePath) {
        try {
            FileInputStream inputStream = new FileInputStream(inputFilePath);
            OutputStream outputStream = new FileOutputStream(outputFilePath);

            convert(inputStream, outputStream);

            inputStream.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public void convert(final InputStream inputStream, final OutputStream outputStream) {
        try {
            com.spire.doc.Document doc = new com.spire.doc.Document();
            doc.loadFromStream(inputStream, FileFormat.Docx);
            ToPdfParameterList ppl = new ToPdfParameterList();
            ppl.isEmbeddedAllFonts(true);
            ppl.setDisableLink(true);
            ppl.setPdfConformanceLevel(PdfConformanceLevel.Pdf_A_3_B);
            doc.setJPEGQuality(40);
            doc.saveToStream(outputStream, FileFormat.PDF);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public String getName() {
        return "Spire Ice (payed) Converter";
    }
}
