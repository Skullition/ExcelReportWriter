package skullition.UI;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.layout.GridPane;
import javafx.stage.FileChooser;
import javafx.stage.Window;

import java.io.File;
import java.net.URL;
import java.util.ResourceBundle;

public class FXMLController implements Initializable {

    @FXML
    private GridPane gridPane;
    @FXML
    private Button loadFileButton;
    @FXML
    private Button makePdfButton;

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        loadFileButton = new Button();
        makePdfButton = new Button();
    }

    @FXML
    private void loadFile(ActionEvent event) {
        FileChooser fileChooser = new FileChooser();
        Window window = gridPane.getScene().getWindow();
        File selectedFile = fileChooser.showOpenDialog(window);

    }
}
