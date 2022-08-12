import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
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
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadExcel {
    public static final Logger LOGGER = LogManager.getLogger();
    /**
     * String of URL location with GreatNusa image
     */
    public static final String GREATNUSA_IMAGE_URL = "https://greatnusa.com/pluginfile.php/1/theme_edumy/headerlogo2/1658542671/Great%20Nusa%20Logo-05_transparen_R.jpg";
    public static final BaseColor DARK_BLUE = new BaseColor(68, 114, 196);
    public static final BaseColor LIGHT_BLUE = new BaseColor(221, 235, 247);

    public static void main(String[] args) {
        new ReadExcel().createPdf(args);
    }

    /**
     * Constructs a PDF file based on data from input
     *
     * @param args the Microsoft Excel document to be read
     */
    private void createPdf(String[] args) {
        try {
            // Read XSSFWorkbook
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(args[0]));
            XSSFSheet sheet = workbook.getSheetAt(0);


            for (Row row : sheet) {
                // Iterate through all column for each row
                if (row.getRowNum() == 0) {
                    continue;
                }
                List<String> cellValues = new ArrayList<>();
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
                    cellValues.add(cellContent);
                }
                CreatePdfBasedOnReportType(cellValues);
            }
        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }

    private void CreatePdfBasedOnReportType(List<String> cellValues) throws DocumentException, IOException {
        String dataType = cellValues.get(0);
        System.out.println(cellValues);
        if (dataType.equalsIgnoreCase("Retail")) {
            createRetail(cellValues);
        } else if (dataType.equalsIgnoreCase("B2B")) {

        } else if (dataType.equalsIgnoreCase("R&B")) {

        } else if (dataType.equalsIgnoreCase("Non BINUS")) {

        } else {
            throw new IllegalArgumentException();
        }
    }

    private void createRetail(List<String> cellValues) throws DocumentException, IOException {
        // Create new Document
        Document document = new Document(PageSize.A4.rotate());
        PdfWriter.getInstance(document, new FileOutputStream(cellValues.get(1)));
        document.open();

        Image image = Image.getInstance(new URL(GREATNUSA_IMAGE_URL));
        image.setAlignment(Image.ALIGN_RIGHT);
        image.scalePercent(5f);
        document.add(image);

        Paragraph headerPhrase = new Paragraph("LAPORAN PENJUALAN KURSUS", new Font(Font.FontFamily.HELVETICA, 16));
        document.add(headerPhrase);
        //TODO: get value for total
        addBodyParagraph(document, cellValues.get(1), cellValues.get(2), cellValues.get(3), cellValues.get(4));
        PdfPTable table = new PdfPTable(8);
        table.addCell(createHeaderCell("Nama Kursus"));
        table.addCell(createHeaderCell("Harga Kursus"));
        table.addCell(createHeaderCell("Jumlah Transaksi"));
        table.addCell(createHeaderCell("Potongan Payment Gateway"));
        table.addCell(createHeaderCell("Persentase"));
        table.addCell(createHeaderCell("Pendapatan sebelum pajak"));
        table.addCell(createHeaderCell("Persentase pajak"));
        table.addCell(createHeaderCell("Pendapatan Akhir"));

        document.add(table);
        document.close();
    }

    private void addBodyParagraph(Document document, String to, String email, String period, String total) throws DocumentException {
        Font font = new Font(Font.FontFamily.HELVETICA, 11);
        document.add(new Paragraph("Kepada: " + to, font));
        document.add(new Paragraph("Email: " + email, font));
        document.add(new Paragraph("Periode: " + period, font));
        document.add(new Paragraph("Total: " + total, font));
    }

    private PdfPCell createHeaderCell(String cellContent) {
        Font darkBlue = new Font(Font.FontFamily.HELVETICA, 10, 0, BaseColor.WHITE);
        PdfPCell cell = new PdfPCell(new Phrase(cellContent, darkBlue));
        cell.setBackgroundColor(DARK_BLUE);
        cell.setBorderColor(DARK_BLUE);
        return cell;
    }
}
