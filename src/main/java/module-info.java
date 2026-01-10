module det014.memorandum_autograder {
    requires javafx.controls;
    requires javafx.fxml;

    opens det014.memorandum_autograder to javafx.fxml;
    exports det014.memorandum_autograder;
}
