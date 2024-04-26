package poc.converters;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.ConversionFeatures;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.convert.out.html.SdtToListSdtTagHandler;
import org.docx4j.convert.out.html.SdtWriter;
import org.docx4j.fonts.BestMatchingMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.local.JodConverter;

public class HtmlConverter implements Converter {

    static boolean save = true;
    static boolean nestLists = true;

    @Override
    public String getName() {
        return null;
    }

    @Override
    public void convert(final String inputFilePath, final String outputFilePath) {
        try {

            FileInputStream inputStream = new FileInputStream(inputFilePath);
            OutputStream outputStream = new FileOutputStream(outputFilePath);

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(inputStream);
            HTMLSettings htmlSettings = Docx4J.createHTMLSettings();

            htmlSettings.setImageDirPath(inputFilePath + "_files");
            htmlSettings.setImageTargetUri(inputFilePath.substring(inputFilePath.lastIndexOf("/")+1)
                    + "_files");
            htmlSettings.setOpcPackage(wordMLPackage);


    	/* CSS reset, see http://itumbcom.blogspot.com.au/2013/06/css-reset-how-complex-it-should-be.html
    	 *
    	 * motivated by vertical space in tables in Firefox and Google Chrome.

	        If you have unwanted vertical space, in Chrome this may be coming from -webkit-margin-before and -webkit-margin-after
	        (in Firefox, margin-top is set to 1em in html.css)

	        Setting margin: 0 on p is enough to fix it.

	        See further http://www.css-101.org/articles/base-styles-sheet-for-webkit-based-browsers/
    	*/
            String userCSS = null;
            if (nestLists) {
                // use browser defaults for ol, ul, li
                userCSS = "html, body, div, span, h1, h2, h3, h4, h5, h6, p, a, img,  table, caption, tbody, tfoot, thead, tr, th, td " +
                        "{ margin: 0; padding: 0; border: 0;}" +
                        "body {line-height: 1;} ";
            } else {
                userCSS = "html, body, div, span, h1, h2, h3, h4, h5, h6, p, a, img,  ol, ul, li, table, caption, tbody, tfoot, thead, tr, th, td " +
                        "{ margin: 0; padding: 0; border: 0;}" +
                        "body {line-height: 1;} ";

            }
            htmlSettings.setUserCSS(userCSS);


            //Other settings (optional)
//    	htmlSettings.setUserBodyTop("<H1>TOP!</H1>");
//    	htmlSettings.setUserBodyTail("<H1>TAIL!</H1>");

            // Sample sdt tag handler (tag handlers insert specific
            // html depending on the contents of an sdt's tag).
            // This will only have an effect if the sdt tag contains
            // the string @class=XXX
//			SdtWriter.registerTagHandler("@class", new TagClass() );

//		SdtWriter.registerTagHandler(Containerization.TAG_BORDERS, new TagSingleBox() );
//		SdtWriter.registerTagHandler(Containerization.TAG_SHADING, new TagSingleBox() );


            // list numbering:  depending on whether you want list numbering hardcoded, or done using <li>.
            if (nestLists) {
                SdtWriter.registerTagHandler("HTML_ELEMENT", new SdtToListSdtTagHandler());
            } else {
                htmlSettings.getFeatures().remove(ConversionFeatures.PP_HTML_COLLECT_LISTS);
            }

            // Refresh the values of DOCPROPERTY fields
            FieldUpdater updater = null;
//		updater = new FieldUpdater(wordMLPackage);
//		updater.update(true);

            // Set up font mapper (optional)
            // We don't add CSS for a font which isn't physically present.
            // TODO: consider web fonts?
//		Mapper fontMapper = new IdentityPlusMapper(); // better for Windows
            Mapper fontMapper = new BestMatchingMapper(); // better for Linux
            wordMLPackage.setFontMapper(fontMapper);

            // .. example of mapping font Times New Roman which doesn't have certain Arabic glyphs
            // eg Glyph "ي" (0x64a, afii57450) not available in font "TimesNewRomanPS-ItalicMT".
            // eg Glyph "ج" (0x62c, afii57420) not available in font "TimesNewRomanPS-ItalicMT".
            // to a font which does
            PhysicalFont font
                    = PhysicalFonts.get("Arial Unicode MS");
            // make sure this is in your regex (if any)!!!
//		if (font!=null) {
//			fontMapper.put("Times New Roman", font);
//			fontMapper.put("Arial", font);
//		}
//		fontMapper.put("Libian SC Regular", PhysicalFonts.get("SimSun"));


            // output to an OutputStream.
            OutputStream os;
            if (save) {
                os = new FileOutputStream(inputFilePath + ".html");
            } else {
                os = new ByteArrayOutputStream();
            }

            // If you want XHTML output
            Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);

            //Don't care what type of exporter you use
//		Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_NONE);
            //Prefer the exporter, that uses a xsl transformation
            Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
            //Prefer the exporter, that doesn't use a xsl transformation (= uses a visitor)
//		Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_EXPORT_PREFER_NONXSL);

            if (save) {
                System.out.println("Saved: " + inputFilePath + ".html ");
            } else {
                System.out.println( ((ByteArrayOutputStream)os).toString() );
            }

            // Clean up, so any ObfuscatedFontPart temp files can be deleted
            if (wordMLPackage.getMainDocumentPart().getFontTablePart()!=null) {
                wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
            }
            // This would also do it, via finalize() methods
            htmlSettings = null;
            wordMLPackage = null;

            inputStream.close();
            outputStream.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
