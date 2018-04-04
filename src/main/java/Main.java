import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;


/*
Переписал программу с испрользованием встраиваемой базы SQLite

запуск [java -cp MonthlyReport.jar Main]

'Файл исходных данных должен содержать:\n'
'Колонка 1(А) - ФИО учасников\n'
'Колонка 2(B) - Филиал\n'
'Колонка 3(С) - Количество человеко-часов\n'
'Ячейка D1(4:1) - Название мероприятия\n'
'Ячейка E1(5:1) - Датаы мероприятия, для формирования списка присутствующих\n\n'
'!!! Обязательно проверять людей добавленных в ИТЦ и Администрацию!!!')
 */

/**
 *
 * @author Sergei Ermenkov
 */

public class Main {

    public static void main(String[] args) throws IOException, SQLException {

        if (args.length > 0) {
            switch (args[0]) {
                case "-i":
                    /*
                    Если семинар на стыке 2х месяцев, то записывать как 2 семинара (1 в одном месяце 2 в другом месяце)
                    Даты для техучебы указывать (начало и конец месяца)
                    A1 - дата начала обучения
                    B1 - дата окончания обучения
                    C1 - любой знак(.)
                    A2 - Тема мероприятия
                    B2 - Приказ на мероприятия
                    C2 - тип мероприятия
                    A3...n  - ФИО
                    B3...n  - должность (филиал)
                    C3...n  - кол-во человекочасов
                     */
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

        //String path = "src/main/resources/Blank.xlsx";
        //String path1 = "/Users/mac/IdeaProjects/MonthlyReport/src/main/resources/Spisok0817.xlsx";

        //String jan = "src/main/resources/01/Spisok_01_18.xlsx";
        //String feb = "src/main/resources/02/Spisok_02_18.xlsx";
        Data data = new Data();
        //data.importExcel(jan);
        //data.importExcel(feb);
        data.getReport(1, 2018);
        //data.getReport(2, 2018);

    }
}
