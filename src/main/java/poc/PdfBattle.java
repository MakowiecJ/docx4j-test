package poc;

import poc.converters.Convertapi;
import poc.converters.Converter;
import poc.converters.FOConverter;
import poc.converters.LibreConverter;
import poc.converters.SpireIceConverter;
import poc.converters.XWPFConverter;

public class PdfBattle {

    public static void main(String[] args) {
        PdfBattle pdfBattle = new PdfBattle();
        FOConverter foConverter = new FOConverter();
        LibreConverter libreConverter = new LibreConverter();
        XWPFConverter xWPFConverter = new XWPFConverter();
        SpireIceConverter spireIveConverter = new SpireIceConverter();
        Convertapi convertapiConverter = new Convertapi();
        final int sampleSize = 7;

        long[] foTimes = pdfBattle.measureConverterTimes(foConverter, sampleSize);
        long[] libreTimes = pdfBattle.measureConverterTimes(libreConverter, sampleSize);
        libreConverter.stopOffice();
        long[] xwpfTimes = pdfBattle.measureConverterTimes(xWPFConverter, sampleSize);
        long[] spireTimes = pdfBattle.measureConverterTimes(spireIveConverter, sampleSize);
        long[] convertapiTimes = pdfBattle.measureConverterTimes(convertapiConverter, sampleSize);


        printResults(foConverter.getName(), foTimes);
        printResults(libreConverter.getName(), libreTimes);
        printResults(xWPFConverter.getName(), xwpfTimes);
        printResults(spireIveConverter.getName(), spireTimes);
        printResults(convertapiConverter.getName(), convertapiTimes);

    }

    private long[] measureConverterTimes(final Converter converter, final int sampleSize) {
        long[] times = new long[sampleSize];

        for (int i = 0; i < sampleSize; i++) {
            long startTime = System.currentTimeMillis();
//            foConverter.convert(inputFile1, outputFile1);
            converter.convert("C:\\Workspace\\test\\docx4j-test\\src\\main\\resources\\docxs\\demo" + i + ".docx",
                    "C:\\Workspace\\test\\docx4j-test\\pdfbattle\\" + converter.getName() + i + ".pdf");
            long endTime = System.currentTimeMillis();
            times[i] = endTime - startTime;

        }

        return times;
    }

    private static void printResults(final String converterName, final long[] results) {
        for (int i = 0; i < results.length; i++) {
            System.out.println(converterName + i + ": " + results[i] + "ms");
        }
    }

}
