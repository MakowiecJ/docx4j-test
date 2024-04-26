package poc.converters;

public interface Converter {

    String getName();
    void convert(String inputFilePath, String outputFilePath);
}
