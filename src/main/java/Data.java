import data.DatePeriud;
import data.Event;
import data.EventTypes;
import eхcel.CellData;
import eхcel.ExcelStyles;
import eхcel.Util;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.*;
import data.Person;
import storage.SQLiteStorage;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

//TODO попробовать написать с помощю org.apache.poi.ss.util.CellUtil.createCell(Row row, int column, java.lang.String value)
class Data {
    private final SQLiteStorage storage = new SQLiteStorage();
    private Map<String, String> branchPatterns;

    void importExcel(String file) throws IOException, SQLException {
        try (FileInputStream fileInputStream = new FileInputStream(file);
             XSSFWorkbook excelBook = new XSSFWorkbook(fileInputStream)) {
            validImportData(excelBook);

            for (Sheet sheet : excelBook) {
                Event event = getEventFromExcel(sheet);
                for (int i = 2; i < sheet.getLastRowNum() + 1; i++) {
                    Row row = sheet.getRow(i);
                    event.addPerson(getPersonFromExcel(row));
                }
                storage.addEventPersons(event);
            }
        } catch (FileNotFoundException e) {
            System.out.printf("Не найден файл (*.xlsx): %s", file);
        }

    }

    private Event getEventFromExcel(Sheet sheet) {
        String eventName = sheet.getRow(1).getCell(0).getStringCellValue();
        String decree = sheet.getRow(1).getCell(1).getStringCellValue();
        DatePeriud date = new DatePeriud(sheet.getRow(0), 0, 1);
        EventTypes type = Enum.valueOf(EventTypes.class, sheet.getRow(1).getCell(2).getStringCellValue());

        return new Event(type, eventName, decree, date);
    }

    private Person getPersonFromExcel(Row row) throws SQLException {
        String name = row.getCell(0).getStringCellValue();
        String position = row.getCell(1).getStringCellValue();
        String branch = findBranch(position);
        int manHours = (int) row.getCell(2).getNumericCellValue();

        return new Person(name, position, branch, manHours);
    }


    private void validImportData(XSSFWorkbook excelBook) {
        StringBuilder errors = new StringBuilder();
        for (Sheet sheet : excelBook) {
            //Использовать foreach нельзя так как в таком случае он пропускает пустые значения
            for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
                Row row = sheet.getRow(i);
                if (row == null ||
                        row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null ||
                        row.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null ||
                        row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null) {
                    errors.append("\nПустое значение ячейки в ")
                            .append(sheet.getSheetName())
                            .append(" : Строка")
                            .append(((row == null) ? 1 : row.getRowNum() + 1));
                }
            }
        }
        if (errors.length() != 0) throw new IllegalArgumentException(errors.toString());
    }

    private String findBranch(String position) throws SQLException {
        if (branchPatterns == null)
            branchPatterns = storage.getBranchPatterns();
        String positionString = position.toLowerCase().replaceAll("\\s", "");
        for (Map.Entry<String, String> entry : branchPatterns.entrySet()) {
            if (positionString.contains(entry.getKey()))
                return entry.getValue();
        }
        //Если не совпало ни с одним из паттернов значит это Администрация
        return "Администрация";
    }

    /*
    загрузить список для генерации листов excel

    цикл генерируем лист
    задаем форматирование

    загружаем шапку листа
    генерируем шапку
    загружаем список для генерации листа
    цикл генерируем данные в лист
    загружаем окончание листа
    генерируем окончание

    запись в файл
    */
    void getReport(int month, int year) throws SQLException, IOException {

        DatePeriud date = new DatePeriud(month, year);

        XSSFWorkbook workbook = new XSSFWorkbook();
        ExcelStyles excelStyles = new ExcelStyles(workbook);

        //Создаем Set всех филиалов (так как они могут повторяться) у которых было Обучение и Тех учеба
        // за указанный период
        Set<String> branches = storage.getBranchesStrings(date, EventTypes.ОБУЧЕНИЕ);
        branches.addAll(storage.getBranchesStrings(date, EventTypes.ТЕХУЧЕБА));

        //Если за указанный периуд мероприятий не проводилось то завершение программы
        if (branches.isEmpty()) {
            throw new NullPointerException("За период с " + date.getBeginDate() + " по " + date.getEndDate() + ". Мероприятий не проводилось.");
        }

        for (String branch : branches) {
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(branch));

            //-------Форматирование листа и добавление шапки-------

            // Установка ширины колонок на листе
            sheet.setColumnWidth(0, 1792);
            sheet.setColumnWidth(1, 9069);
            sheet.setColumnWidth(2, 7350);
            sheet.setColumnWidth(3, 4717);
            // Обеденение ячеек
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 2));
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 3));
            sheet.addMergedRegion(new CellRangeAddress(6, 6, 0, 3));
            // Добавление шапки в лист
            List<CellData> reportHeader = storage.getTemplate("report_header");
            for (CellData cellData : reportHeader) {
                cellData.addCell(sheet, excelStyles);
            }
            //Подстановка в шапку переменных полей
            //Название филиала
            if ("Администрация".equals(branch)) {
                sheet.getRow(3).getCell(0).setCellValue("В Администрация OOO «Газпром трансгаз Москва»");
            } else {
                sheet.getRow(3).getCell(0).setCellValue("В филиал " + branch);
            }
            //Дата
            //todo исправить вывод месяца на вывод с маленькой буквы
            sheet.getRow(6).getCell(0)
                    .setCellValue("за " + date.getBeginDate().format(DateTimeFormatter.ofPattern("LLLL YYYY", Locale.forLanguageTag("ru"))) + " года");


            //-------Добавление основных данных в лист(людей)-------

            int cursorRowNum = 10;

            List<Event> events = storage.getEvents(date, EventTypes.ОБУЧЕНИЕ, branch);
            events.addAll(storage.getEvents(date, EventTypes.ТЕХУЧЕБА, branch));

            for (Event event : events) {
                int startRowMerged = cursorRowNum;
                new CellData(cursorRowNum, 2, event.getName(), "t12alCCWB").addCell(sheet, excelStyles);

                for (Person person : event.getPersons()) {
                    new CellData(cursorRowNum, 0, cursorRowNum - 9, "t12alCCWB").addCell(sheet, excelStyles);
                    new CellData(cursorRowNum, 1, person.getName(), "t12alLCWB").addCell(sheet, excelStyles);
                    new CellData(cursorRowNum, 3, person.getManHours(), "t12alCCWB").addCell(sheet, excelStyles);
                    cursorRowNum++;
                }

                if (event.personsSize() > 1) {
                    //вырравнивание высоты обединенных ячеек
                    //Вычисление числа строк которое займет название мероприятия после объединения
                    XSSFCellStyle style = excelStyles.getStyle("t12alCCWB");
                    int rowsHigh = Util.getHigh(style.getFont().getFontName(), style.getFont().getFontHeightInPoints(), event.getName(), sheet.getColumnWidthInPixels(2));
                    int rowsIs = cursorRowNum - startRowMerged;
                    float newHigh = (sheet.getDefaultRowHeightInPoints() * rowsHigh) / rowsIs;
                    if (rowsHigh > rowsIs) {
                        for (int i = startRowMerged; i < cursorRowNum; i++) {
                            sheet.getRow(i).setHeightInPoints(newHigh);
                        }
                    }
                    //обединение клеток с одним мероприятием
                    sheet.addMergedRegion(new CellRangeAddress(startRowMerged, cursorRowNum - 1, 2, 2));
                }
            }

            //-------Добавление Итога и хвоста страницы-------

            // Обединение ячеек итога
            sheet.addMergedRegion(new CellRangeAddress(cursorRowNum, cursorRowNum, 0, 2));
            // Добавление хвоста в лист
            List<CellData> reportFooter = storage.getTemplate("report_footer");
            for (CellData cellData : reportFooter) {
                cellData.addCell(sheet, excelStyles, cursorRowNum);
            }
            //Подстановка в формулы в итог
            sheet.getRow(cursorRowNum).getCell(3).setCellFormula("SUM(D11:D" + cursorRowNum + ")");

            //Добавляем свойство "Разместить не более чем на 1 странице"
            XSSFPrintSetup printSetup = sheet.getPrintSetup();
            sheet.setFitToPage(true);
            printSetup.setFitHeight((short) 1);
            printSetup.setFitWidth((short) 1);
        }

        //Запись в файл
        try (FileOutputStream fileOut = new FileOutputStream("Отчет по обучению_" + month + "_" + year + ".xlsx")) {
            workbook.write(fileOut);
        }
    }


    //-------------------------------------------------------------------------------------------------------------


    void getSpisok(int month, int year) throws SQLException, IOException {

        DatePeriud date = new DatePeriud(month, year);

        XSSFWorkbook workbook = new XSSFWorkbook();
        ExcelStyles excelStyles = new ExcelStyles(workbook);

        List<Event> events = storage.getEvents(date, EventTypes.СЕМИНАР);
        events.addAll(storage.getEvents(date, EventTypes.ПРОЧЕЕ));

        //Если за указанный периуд мероприяий не проводилось то завершение программы
        //изменение в проверки условия
        if (events.isEmpty()) {
            throw new NullPointerException("За период с " + date.getBeginDate() + " по " + date.getEndDate() + ". Мероприятий не проводилось.");
        }

        //изменение в переменных
        //todo добавить проверку на совпадение нового имени с именем в книги
        for (Event event : events) {
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(event.getName()));

            //-------Форматирование листа и добавление шапки-------

            // Установка ширины колонок на листе
            sheet.setColumnWidth(0, 1792);
            sheet.setColumnWidth(1, 9069);
            sheet.setColumnWidth(2, 7350);
            sheet.setColumnWidth(3, 4717);

            //изменение ------------------>
            // Обеденение ячеек
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 3));
            sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 3));
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 3));
            //<----------------------------

            // Добавление шапки в лист
            //изменения в вызове
            List<CellData> listHeader = storage.getTemplate("list_header");
            for (CellData cellData : listHeader) {
                cellData.addCell(sheet, excelStyles);
            }

            //изменение ------------------>
            //Подстановка в шапку переменных полей
            //Название мероприятия (выровнять высоту строки)
            //todo заменить ручное указание стиля
            XSSFCellStyle style = excelStyles.getStyle("t12balCW");
            sheet.getRow(4).getCell(0)
                    .setCellValue("участников мероприятия: " + event.getName());
            int rowsHigh = Util.getHigh(style.getFont().getFontName(),
                    style.getFont().getFontHeightInPoints(),
                    event.getName(),
                    sheet.getColumnWidthInPixels(0) +
                            sheet.getColumnWidthInPixels(1) +
                            sheet.getColumnWidthInPixels(2) +
                            sheet.getColumnWidthInPixels(3));
            if (rowsHigh > 1)
                sheet.getRow(4).setHeightInPoints(sheet.getDefaultRowHeightInPoints() * rowsHigh);

            //Период
            sheet.getRow(5).getCell(0)
                    .setCellValue("в период " + event.getDate());

            //-------Добавление основных данных в лист(людей)-------

            int cursorRowNum = 9;

            for (Person person : event.getPersons()) {
                new CellData(cursorRowNum, 0, cursorRowNum - 8, "t12alCCWB").addCell(sheet, excelStyles);
                new CellData(cursorRowNum, 1, person.getName(), "t12alLCWB").addCell(sheet, excelStyles);
                new CellData(cursorRowNum, 2, person.getPosition(), "t12alCCWB").addCell(sheet, excelStyles);
                new CellData(cursorRowNum, 3, person.getManHours(), "t12alCCWB").addCell(sheet, excelStyles);
                cursorRowNum++;
            }

            //-------Добавление Итога и хвоста страницы-------

            // Обединение ячеек итога
            sheet.addMergedRegion(new CellRangeAddress(cursorRowNum, cursorRowNum, 0, 2));
            // Добавление хвоста в лист
            List<CellData> reportFooter = storage.getTemplate("report_footer");
            for (CellData cellData : reportFooter) {
                cellData.addCell(sheet, excelStyles, cursorRowNum);
            }
            //Подстановка в формулы в итог
            sheet.getRow(cursorRowNum).getCell(3).setCellFormula("SUM(D10:D" + cursorRowNum + ")");

            //Добавляем свойство "Разместить не более чем на 1 странице"
            XSSFPrintSetup printSetup = sheet.getPrintSetup();
            sheet.setFitToPage(true);
            //todo убран ограничитель на 1 страницу по высоте
            printSetup.setFitWidth((short) 1);
        }

        //Запись в файл
        //todo изменен вывод файла
        try (FileOutputStream fileOut = new FileOutputStream("Список участников семинаров_" + month + "_" + year + ".xlsx")) {
            workbook.write(fileOut);
        }
    }
}