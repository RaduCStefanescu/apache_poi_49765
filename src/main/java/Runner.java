import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPicture;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Optional;

public class Runner {

    private static final String INPUT_DOCUMENT = "input.docx";
    private static final String IMAGE_FILE = "signature.jpg";
    private static final String OUTPUT_DOCUMENT = "output.docx";

    public static void main(String[] args) throws Exception{
        Path inputPath = Paths.get(Paths.get(ClassLoader.getSystemResource(INPUT_DOCUMENT).toURI()).toString());
        Path imagePath = Paths.get(Paths.get(ClassLoader.getSystemResource(IMAGE_FILE).toURI()).toString());
        Path outputPath = Paths.get(Paths.get(ClassLoader.getSystemResource(OUTPUT_DOCUMENT).toURI()).toString());

        //load the input document
        Optional<XWPFDocument> document = DocxHandler.load(inputPath);
        Optional<XWPFPicture> picture = Optional.empty();

        if(document.isPresent()){
            try {
                //add the picture to the document
                picture = DocxHandler.addPicture(document.get(), imagePath);
            }catch(Exception exception){
                System.out.println("Unable to add picture to input Document!");
                exception.printStackTrace();
                return ;
            }

            if(!picture.isPresent()){
                System.out.println("Added picture is null, something went wrong!");
                return ;
            }

            try {
                //write the document containing the picture to the output file
                DocxHandler.write(outputPath, document.get());
            }catch(Exception exception){
                System.out.println("Unable to write output document!");
                exception.printStackTrace();
            }
        }

    }
}
