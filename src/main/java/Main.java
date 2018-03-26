import eхcel.CellData;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.sql.SQLException;
import java.util.List;
import java.util.Map;
import java.util.logging.ConsoleHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;

/*
Переписал программу с испрользованием встраиваемой базы SQLite

запуск [java -cp MonthlyReport.jar Main]
 */

/**
 * @author Sergei Ermenkov
 */

//TODO Подумать как обрабатывть тех учебу (ей нужен отдельный тип)
    //TODO в базе данных убрать ограничения type у мероприятия
public class Main {

    public static void main(String[] args) throws IOException, SQLException {

        if (args.length > 0) {
            switch (args[0]) {
                case "-i":
                    new Data().importExcel(args[1]);
                    break;
                case "-r":
                    try {
                        new Data().getReport(Integer.valueOf(args[1]), Integer.valueOf(args[2]));
                    } catch (NumberFormatException e){
                        System.out.println("Месяц и год должны быть числа");
                    }
            }
        } else {
            System.out.println("Используйте:");
            System.out.println("-i ИМЯ_ФАЙЛА.xlsx       Импортировать данные из файла");
            System.out.println("-r МЕСЯЦ ГОД            Сформировать отчет");
        }

        //String path = "/Users/mac/IdeaProjects/MonthlyReport/src/main/resources/Blank.xlsx";
        //String path1 = "/Users/mac/IdeaProjects/MonthlyReport/src/main/resources/Spisok0817.xlsx";
        //Data data = new Data();
        //data.importExcel(path);
        //data.getReport(2, 2018);
    }
}
