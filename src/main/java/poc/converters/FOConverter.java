package poc.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.FopFactoryBuilder;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.fo.renderers.FORendererApacheFOP;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class FOConverter implements Converter {

    @Override
    public void convert(final String inputFilePath, final String outputFilePath) {
        try {
            FileInputStream inputStream = new FileInputStream(inputFilePath);
            OutputStream outputStream = new FileOutputStream(outputFilePath);

            WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(inputStream);

            // use FO converter
            FOSettings foSettings = new FOSettings(mlPackage);
            FopFactoryBuilder fopFactoryBuilder = FORendererApacheFOP.getFopFactoryBuilder(foSettings);
            FopFactory fopFactory = fopFactoryBuilder.build();

            FOUserAgent foUserAgent = FORendererApacheFOP.getFOUserAgent(foSettings, fopFactory);
            foUserAgent.setTitle("my title");
            foUserAgent.getRendererOptions().put("version", "2.0");

            Docx4J.toFO(foSettings, outputStream, Docx4J.FLAG_EXPORT_PREFER_XSL); // USE Docx4J.FLAG_EXPORT_PREFER_NONXSL (less demanding)

            inputStream.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public String getName() {
        return "Apache FO Converter";
    }
}
