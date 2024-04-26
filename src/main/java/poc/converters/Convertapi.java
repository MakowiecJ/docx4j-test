package poc.converters;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.FopFactory;
import org.apache.fop.apps.FopFactoryBuilder;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.fo.renderers.FORendererApacheFOP;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import com.convertapi.client.Config;
import com.convertapi.client.ConvertApi;
import com.convertapi.client.Param;

public class Convertapi implements Converter {
    @Override
    public String getName() {
        return "Convertapi (cloud)(payed)";
    }

    @Override
    public void convert(final String inputFilePath, final String outputFilePath) {
        try {
//            FileInputStream inputStream = new FileInputStream(inputFilePath);
//            OutputStream outputStream = new FileOutputStream(outputFilePath);

            // Code snippet is using the ConvertAPI Java Client: https://github.com/ConvertAPI/convertapi-java

            Config.setDefaultSecret("YbpJDTk4hoQiFX6d");
            ConvertApi.convert("docx", "pdf",
                    new Param("File", Paths.get(inputFilePath))
            ).get().saveFile(Path.of(outputFilePath));
//                    .saveFilesSync(Paths.get(outputFilePath));

//            inputStream.close();
//            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
