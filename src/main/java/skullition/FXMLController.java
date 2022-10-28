package skullition;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.GridPane;
import javafx.stage.FileChooser;
import javafx.stage.Window;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

public class FXMLController implements Initializable {

    @FXML
    private GridPane gridPane;
    @FXML
    private Button loadFileButton;
    @FXML
    private Button makePdfButton;
    @FXML
    private Label fileNameLabel;
    @FXML
    private TableView<Data> tableView;
    private File chosenFile;

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        makePdfButton.setDisable(true);
    }

    @FXML
    private void loadFile(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        Window window = gridPane.getScene().getWindow();
        this.chosenFile = fileChooser.showOpenDialog(window);


        fileNameLabel.setText(chosenFile.getName());
        if (chosenFile.getName().endsWith("xlsx")) {
            makePdfButton.setDisable(false);
            updateTableView(chosenFile);
        }

    }

    private void updateTableView(File file) {
        TableColumn<Data, String> reportTypeColumn = new TableColumn<>("Tipe Report");
        TableColumn<Data, String> personNameColumn = new TableColumn<>("Nama");
        TableColumn<Data, String> emailColumn = new TableColumn<>("Email");
        TableColumn<Data, String> periodColumn = new TableColumn<>("Periode");
        TableColumn<Data, String> courseNameColumn = new TableColumn<>("Nama Kursus");
        TableColumn<Data, String> coursePriceColumn = new TableColumn<>("Harga Kursus");
        TableColumn<Data, String> transactionAmountColumn = new TableColumn<>("Jumlah Transaksi");
        TableColumn<Data, String> totalPaymentGatewayColumn = new TableColumn<>("Total Biaya Payment Gateway");
        TableColumn<Data, String> cutPaymentGatewayColumn = new TableColumn<>("Potongan Payment Gateway");
        TableColumn<Data, String> administrationColumn = new TableColumn<>("Biaya Administrasi");
        TableColumn<Data, String> incomePercentageColumn = new TableColumn<>("Persentase Pendapatan");
        TableColumn<Data, String> incomeBeforeTaxColumn = new TableColumn<>("Pendapatan Sebelum Pajak");
        TableColumn<Data, String> taxPercentageColumn = new TableColumn<>("Persentase Pajak");
        TableColumn<Data, String> taxAmountColumn = new TableColumn<>("Nominal Pajak");
        TableColumn<Data, String> endIncomeColumn = new TableColumn<>("Pendapatan Akhir");

        reportTypeColumn.setCellValueFactory(new PropertyValueFactory<>("reportType"));
        personNameColumn.setCellValueFactory(new PropertyValueFactory<>("personName"));
        emailColumn.setCellValueFactory(new PropertyValueFactory<>("email"));
        periodColumn.setCellValueFactory(new PropertyValueFactory<>("period"));
        courseNameColumn.setCellValueFactory(new PropertyValueFactory<>("courseName"));
        coursePriceColumn.setCellValueFactory(new PropertyValueFactory<>("coursePrice"));
        transactionAmountColumn.setCellValueFactory(new PropertyValueFactory<>("transactionAmount"));
        totalPaymentGatewayColumn.setCellValueFactory(new PropertyValueFactory<>("totalPaymentGateway"));
        cutPaymentGatewayColumn.setCellValueFactory(new PropertyValueFactory<>("cutPaymentGateway"));
        administrationColumn.setCellValueFactory(new PropertyValueFactory<>("administration"));
        incomePercentageColumn.setCellValueFactory(new PropertyValueFactory<>("incomePercentage"));
        incomeBeforeTaxColumn.setCellValueFactory(new PropertyValueFactory<>("incomeBeforeTax"));
        taxPercentageColumn.setCellValueFactory(new PropertyValueFactory<>("taxPercentage"));
        taxAmountColumn.setCellValueFactory(new PropertyValueFactory<>("taxAmount"));
        endIncomeColumn.setCellValueFactory(new PropertyValueFactory<>("endIncome"));

        // Iterate through sheet
        XSSFWorkbook workbook;
        try {
            workbook = new XSSFWorkbook(file);
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
        XSSFSheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            List<String> cellValues = new ArrayList<>();

            for (Cell cell : row) {
                String cellContent = ReadExcel.getValueFromCell(cell);
                cellValues.add(cellContent);
            }
            addTableFromArray(cellValues);
        }


        tableView.getColumns().setAll(reportTypeColumn, personNameColumn, emailColumn, periodColumn, courseNameColumn, coursePriceColumn, transactionAmountColumn, totalPaymentGatewayColumn, cutPaymentGatewayColumn, administrationColumn, incomePercentageColumn, incomeBeforeTaxColumn, taxPercentageColumn, taxAmountColumn, endIncomeColumn);
        try {
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void addTableFromArray(List<String> cellValues) {
        Data data = new Data(cellValues.get(0), cellValues.get(1), cellValues.get(2), cellValues.get(3), cellValues.get(4), cellValues.get(5), cellValues.get(6), cellValues.get(7), cellValues.get(8), cellValues.get(9), cellValues.get(10), cellValues.get(11), cellValues.get(12), cellValues.get(13), cellValues.get(14));
        tableView.getItems().add(data);
    }

    @FXML
    private void makePdf(ActionEvent event) {
        ReadExcel readExcel = new ReadExcel();
        readExcel.createPdf(chosenFile);
    }
}
