package skullition;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;

import java.net.URL;
import java.util.ResourceBundle;

public class FXMLController implements Initializable {

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
        System.out.println("lsdflsdl");
    }
}
