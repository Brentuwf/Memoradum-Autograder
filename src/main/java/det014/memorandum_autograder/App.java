package det014.memorandum_autograder;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

import acp.project2.SpellCheckerController;

/**
 * JavaFX App
 */
public class App extends Application {
	public static final int WINDOW_HIGTH = 400;
	public static final int WINDOW_WIDTH = 600;
	public static final String FXML_CONTROLLER_FILE = "/autograder.fxml";

    @Override
    public void start(Stage stage) throws IOException {
    	FXMLLoader loader = new FXMLLoader(getClass().getResource(FXML_CONTROLLER_FILE));
		Parent root = loader.load();
		
		AutograderController controller = loader.getController();
	        
		Scene scene = new Scene(root, WINDOW_WIDTH, WINDOW_HIGTH);
		stage.setScene(scene);
		stage.setTitle("Memorandum Autograder");
		stage.show();
    }

    public static void main(String[] args) {
        launch();
    }

}