import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Iterator;

public class ReadExcel {
    public static final Logger LOGGER = LogManager.getLogger();
    public static final String GREATNUSA_IMAGE_URL = "https://greatnusa.com/pluginfile.php/1/theme_edumy/headerlogo2/1658542671/Great%20Nusa%20Logo-05_transparen_R.jpg";

    public static void main(String[] args) {
        createPdf(args);
    }

    private static void createPdf(String[] args) {
        try {
            // Read XSSFWorkbook
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(args[0]));
            XSSFSheet sheet = workbook.getSheetAt(0);


            for (Row row : sheet) {
                // Iterate through all column for each row
                if (row.getRowNum() == 0) {
                    continue;
                }
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellContent;
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            cellContent = String.valueOf(cell.getNumericCellValue());
                            break;
                        case STRING:
                            cellContent = cell.getStringCellValue();
                            break;
                        default:
                            throw new IllegalStateException("Unexpected value: " + cell.getCellType());
                    }

                    Document document;
                    if (cell.getAddress().getColumn() == 0) {
                        //
                        if (cellContent.equalsIgnoreCase("Retail")) {
                            createRetail();
                        } else if (cellContent.equalsIgnoreCase("B2B")) {

                        } else if (cellContent.equalsIgnoreCase("R&B")) {

                        } else if (cellContent.equalsIgnoreCase("Non BINUS")) {

                        } else {
                            throw new IllegalArgumentException();
                        }
                    }


                }
            }
            // Close Document object
//            document.close();

        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }

    private static void createRetail() throws DocumentException, IOException {
        // Create new Document
        Document document = new Document(PageSize.A4.rotate());
        PdfWriter.getInstance(document, new FileOutputStream("OutputReportAuthor.pdf"));
        document.open();

        Image image = Image.getInstance(new URL(GREATNUSA_IMAGE_URL));
        image.setAlignment(Image.ALIGN_RIGHT);
        image.scalePercent(5f);
//        image.setAbsolutePosition(36, 400);
        document.add(image);

        Paragraph headerPhrase = new Paragraph("LAPORAN PENJUALAN KURSUS", new Font(Font.FontFamily.HELVETICA, 16));
        document.add(headerPhrase);

        document.close();
    }
}
