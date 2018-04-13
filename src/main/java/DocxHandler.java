
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;

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
                        300, 300));
            }
        }

        return Optional.empty();
    }

    private static void clearRuns(XWPFParagraph paragraph){
        int runsNo = paragraph.getRuns().size();
        for(int index = 0; index < runsNo; ++index){
            paragraph.removeRun(0);
        }
    }
}
