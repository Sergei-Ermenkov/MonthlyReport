import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.ConsoleHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;



/**
 * @author Sergei Ermenkov
 */
public class Test {
    private static final Logger LOGGER = Logger.getLogger(Test.class.getName());

    public static void main(String[] args) throws IOException {
        // Настройки протоколирования
        Handler handler = new ConsoleHandler();
        handler.setLevel(Level.FINEST);
        LOGGER.setLevel(Level.FINEST);
        LOGGER.setUseParentHandlers(false);
        LOGGER.addHandler(handler);

        String path = "/Users/mac/IdeaProjects/MonthlyReport/src/main/java/Blank.xlsx";
        String path1 = "/Users/mac/IdeaProjects/MonthlyReport/src/main/java/2017.xlsx";
//        LOGGER.log(Level.FINEST, "Путь: {0}", path);

        XSSFWorkbook workbook = new XSSFWorkbook();

        ExcelData e = new ExcelData();
        e.readFromExcel(path);
        e.makeReport();
        e.writeToExcel("workbook.xls");
//        new Test().extractor(path);
    }

}
