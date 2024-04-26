package poc.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;


public class XWPFConverter implements Converter {

    @Override
    public void convert(final String inputFilePath, final String outputFilePath) {
        try {
            FileInputStream inputStream = new FileInputStream(inputFilePath);
            OutputStream outputStream = new FileOutputStream(outputFilePath);

            XWPFDocument document = new XWPFDocument(inputStream);

            PdfOptions options = PdfOptions.create();

            org.apache.poi.xwpf.converter.pdf.PdfConverter.getInstance().convert(document, outputStream, options);

            inputStream.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public String getName() {
        return "XWPF Converter";
    }
}
