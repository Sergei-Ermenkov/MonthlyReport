package zimenki.javafxControllers;

import zimenki.Main;
import zimenki.storage.ImportToStorage;
import zimenki.eхcel.NotIntersectDateException;
import zimenki.eхcel.WrongImportFilenameException;

import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.stage.FileChooser;

import java.io.File;
import java.io.IOException;
import java.sql.SQLException;

public class RootLayoutController {

    private Main main;

    public void setMain(Main main) {
        this.main = main;
    }


    @FXML
    private void handleOpen() throws IOException, SQLException {
        FileChooser fileChooser = new FileChooser();
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter(
                "Excel files (*.xlsx)", "*.xlsx");
        fileChooser.getExtensionFilters().add(extFilter);
        File file = fileChooser.showOpenDialog(main.getPrimaryStage());

        if (file != null) {
            try {
                new ImportToStorage().importFromExcel(file);
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.initOwner(main.getPrimaryStage());
                alert.setTitle(null);
                alert.setHeaderText(null);
                alert.setContentText("Загрузка данных завершена");
                alert.showAndWait();
            } catch (WrongImportFilenameException |
                    NotIntersectDateException e){
                Alert alert = new Alert(Alert.AlertType.ERROR);
                alert.initOwner(main.getPrimaryStage());
                alert.setTitle("Ошибка");
                alert.setHeaderText(null);
                alert.setContentText(e.getMessage());
                alert.showAndWait();
            }

        }
    }


    @FXML
    private void handleHelpFormat() {
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.initOwner(main.getPrimaryStage());
        alert.setTitle("Предупреждение");
        alert.setHeaderText(null);
        alert.setContentText("Файл должен иметь вид (Spisok_03_2018.xlsx) обязательно должен быть указан месяц\n" +
                "Если семинар на стыке 2х месяцев, то записывать как 2 семинара (1 в одном месяце 2 в другом месяце)\n" +
                "Даты для техучебы указывать (начало и конец месяца)\n" +
                "  A1 - дата начала обучения\n" +
                "  B1 - дата окончания обучения\n" +
                "  C1 - любой знак(.)\n" +
                "  A2 - Тема мероприятия\n" +
                "  B2 - Приказ на мероприятия\n" +
                "  C2 - тип мероприятия\n" +
                "  A3...n  - ФИО\n" +
                "  B3...n  - должность (филиал)\n" +
                "  C3...n  - кол-во человекочасов");
        alert.showAndWait();
    }

    @FXML
    private void handleExit() {
        System.exit(0);
    }


}


