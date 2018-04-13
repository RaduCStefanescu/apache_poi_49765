
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;

import static org.apache.poi.util.Units.EMU_PER_PIXEL;

public class DocxHandler {

    static Optional<XWPFDocument> load(Path originalPath){
        try {
            return Optional.ofNullable(new XWPFDocument(new FileInputStream(originalPath.toFile())));
        }catch(IOException exception){
            return Optional.empty();
        }
    }

    static void write(Path writePath, XWPFDocument document)
            throws IOException, SecurityException {
        document.write(new FileOutputStream(writePath.toFile()));
    }

    static Optional<XWPFPicture> addPicture(XWPFDocument document, Path imagePath)
            throws IOException, InvalidFormatException {

        for(XWPFParagraph paragraph: document.getParagraphs()){
            if(StringUtils.equalsIgnoreCase(paragraph.getText(), "Signature")){
                clearRuns(paragraph);
                XWPFRun newRun = paragraph.createRun();
                return Optional.ofNullable(newRun.addPicture(Files.newInputStream(imagePath),
                        Document.PICTURE_TYPE_JPEG,
                        imagePath.getFileName().toString(),
                        100 * EMU_PER_PIXEL, 100 * EMU_PER_PIXEL));
            }
        }

        return Optional.empty();
    }

    static void addPictureWithFix(XWPFDocument document, Path imagePath)
            throws IOException, InvalidFormatException {

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            if (StringUtils.equalsIgnoreCase(paragraph.getText(), "Signature")) {
                clearRuns(paragraph);
                String blipId = document.addPictureData(Files.newInputStream(imagePath), Document.PICTURE_TYPE_JPEG);
                createPicture(document, paragraph, blipId, 300, 300);
            }
        }
    }

    private static void clearRuns(XWPFParagraph paragraph){
        int runsNo = paragraph.getRuns().size();
        for(int index = 0; index < runsNo; ++index){
            paragraph.removeRun(0);
        }
    }

    private static void createPicture(XWPFDocument document, XWPFParagraph paragraph, String blipId, int width, int height){
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        CTInline inline = paragraph.createRun().getCTR().addNewDrawing().addNewInline();

        String picXml = "" +
                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "         <pic:nvPicPr>" +
                "            <pic:cNvPr id=\"" + 0 + "\" name=\"Generated\"/>" +
                "            <pic:cNvPicPr/>" +
                "         </pic:nvPicPr>" +
                "         <pic:blipFill>" +
                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "            <a:stretch>" +
                "               <a:fillRect/>" +
                "            </a:stretch>" +
                "         </pic:blipFill>" +
                "         <pic:spPr>" +
                "            <a:xfrm>" +
                "               <a:off x=\"0\" y=\"0\"/>" +
                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                "            </a:xfrm>" +
                "            <a:prstGeom prst=\"rect\">" +
                "               <a:avLst/>" +
                "            </a:prstGeom>" +
                "         </pic:spPr>" +
                "      </pic:pic>" +
                "   </a:graphicData>" +
                "</a:graphic>";

        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();
        XmlToken xmlToken = null;
        try
        {
            xmlToken = XmlToken.Factory.parse(picXml);
        }
        catch(XmlException xe)
        {
            xe.printStackTrace();
        }

        inline.set(xmlToken);
        //graphicData.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(0);
        docPr.setName("Picture " + 0);
        docPr.setDescr("Generated");
    }
}
