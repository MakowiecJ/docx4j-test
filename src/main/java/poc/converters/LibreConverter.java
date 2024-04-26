package poc.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.JodConverter;
import org.jodconverter.local.office.LocalOfficeManager;

public class LibreConverter implements Converter {

    private static final LocalOfficeManager officeManager;

    static {
        officeManager = LocalOfficeManager.builder().install().build();
        try {
            officeManager.start();
        } catch (OfficeException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void convert(final String inputFilePath, final String outputFilePath) {
        try {

            FileInputStream inputStream = new FileInputStream(inputFilePath);
            OutputStream outputStream = new FileOutputStream(outputFilePath);

            // Perform the conversion
            JodConverter
                    .convert(inputStream)
                    .as(DefaultDocumentFormatRegistry.DOCX)
                    .to(outputStream)
                    .as(DefaultDocumentFormatRegistry.PDF)
                    .execute();

            inputStream.close();
            outputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public String getName() {
        return "JodConverter (Libre Office)";
    }

    public void stopOffice() {
        try {
            officeManager.stop();
        } catch (OfficeException e) {
            throw new RuntimeException(e);
        }
    }
}
