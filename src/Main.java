
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;

public class Main {

    private static final String SOURCE_DOC_PATH = "/home/user/test_please_delete_me.docx";
    private static final String RESULT_DOC_PATH = "/home/user/test_result_please_delete_me.docx";
    private static final String RESULT_PDF_PATH = "/home/user/test_result_please_delete_me.pdf";

    private static HashMap<String, String> PARAMS = new HashMap<>();

    public static void main(String[] args) throws IOException, InvalidFormatException {
        initMap();
        wordDocProcessor(
                SOURCE_DOC_PATH,
                RESULT_DOC_PATH
        );

        ConvertToPDF(RESULT_DOC_PATH, RESULT_PDF_PATH);

        System.out.println("done");
    }

    public static void ConvertToPDF(String docPath, String pdfPath) {
        try {
            InputStream doc = new FileInputStream(new File(docPath));
            XWPFDocument document = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(pdfPath));
            PdfConverter.getInstance().convert(document, out, options);
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
    }

    public static void wordDocProcessor(String inputFilePath, String outputFilePath) throws IOException,
            InvalidFormatException {
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(inputFilePath));


        doc.getBodyElements().forEach(el -> {
                    if (el instanceof XWPFParagraph) {
                        ((XWPFParagraph) el).getRuns().forEach(f -> {

                                    String currentString = f.getText(0);
                                    if (currentString != null) {
                                        PARAMS.forEach((key, value) -> {

                                            if (currentString.contains(getFullKey(key))) {
                                                f.setText(currentString.replaceAll(getRegex(key), value), 0);
                                            }
                                        });
                                    }
                                }
                        );
                    }
                }
        );

        doc.write(new FileOutputStream(outputFilePath));
    }

    public static void initMap() {
        PARAMS.put("currentDate", LocalDate.now().format(DateTimeFormatter.ofPattern("dd MMM yyyy")));
        PARAMS.put("currentDateTime", LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd MMM yyyy, HH:mm")));
    }

    private static String getRegex(String key) {
        return "\\%P\\{".concat(key).concat("\\}");
    }

    private static String getFullKey(String key) {
        return "%P{".concat(key).concat("}");
    }
}
