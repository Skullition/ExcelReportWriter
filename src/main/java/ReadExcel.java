import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfDocument;
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
import java.util.Iterator;

public class ReadExcel {
    public static final Logger LOGGER = LogManager.getLogger();

    public static void main(String[] args) {
        createPdf(args);
    }

    private static void createPdf(String[] args) {
        try {
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
                            // below for testing
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case STRING:
                            cellContent = cell.getStringCellValue();
                            // below for testing
                            System.out.println(cell.getStringCellValue());
                            break;
                        default:
                            throw new IllegalStateException("Unexpected value: " + cell.getCellType());
                    }

                    Document document = new Document();
                    PdfWriter.getInstance(document, new FileOutputStream("TestOutput.pdf"));
                    document.open();
                    if (cell.getAddress().getColumn() == 0) {
                        //
                    }


                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        }
    }
}
