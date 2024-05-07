package poc.converters;

import java.io.InputStream;
import java.io.OutputStream;

public interface Converter {

    String getName();
    void convert(String inputFilePath, String outputFilePath);
    void convert(InputStream inputStream, OutputStream outputStream);
}
