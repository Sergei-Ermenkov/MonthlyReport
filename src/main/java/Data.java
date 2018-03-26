import eхcel.CellData;
import eхcel.ExcelStyles;
import javafx.print.Collation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.List;
import java.util.Map;


class Data {
    private SQLiteStorage storage = new SQLiteStorage();
    private Map<String, Integer> branchPatterns;

    void importExcel(String file) throws IOException, SQLException {
        // Загрузка книги Excel
        try (FileInputStream fileInputStream = new FileInputStream(file);
             XSSFWorkbook excelBook = new XSSFWorkbook(fileInputStream)) {

            // Проверка на валидность входных данных в excel
            validImportData(excelBook);

            // Проход по всем листам, строкам и полям.
            for (Sheet sheet : excelBook) {
                String eventName;
                String decree;
                LocalDate beginDate = null;
                LocalDate endDate = null;
                EventTypes type;
                int event_id = 0;

                for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
                    // Переменные относящиеся к человеку (строки Excel)
                    String name;
                    String position;
                    int manHours;
                    int branch_id;
                    int person_id;

                    Row row = sheet.getRow(i);
                    // различная обработка строк
                    switch (i) {
                        case 0:
                            beginDate = row.getCell(0).getDateCellValue()
                                    .toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                            endDate = row.getCell(1).getDateCellValue()
                                    .toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                            break;
                        case 1:
                            eventName = row.getCell(0).getStringCellValue();
                            decree = row.getCell(1).getStringCellValue();
                            type = Enum.valueOf(EventTypes.class, row.getCell(2).getStringCellValue());

                            event_id = storage.addEvent(type, eventName, decree, beginDate, endDate);
                            break;
                        default:
                            name = row.getCell(0).getStringCellValue();
                            position = row.getCell(1).getStringCellValue();
                            branch_id = findBranch(position);
                            manHours = (int) row.getCell(2).getNumericCellValue();

                            person_id = storage.addPerson(name, position, branch_id);
                            storage.addPersonToEvent(event_id, person_id, manHours);
                    }
                }
            }
        } catch (FileNotFoundException e){
            System.out.printf("Не найден файл (*.xlsx): %s",file);
        }
    }

    //Использовать foreach нельзя так как в таком случае он пропускает пустые значения
    private void validImportData(XSSFWorkbook excelBook) {
        StringBuilder errors = new StringBuilder();
        for (Sheet sheet : excelBook) {
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

    // Анализирует строку должности и возвращает ID номер филиала из базы
    private int findBranch(String position) throws SQLException {
        if (branchPatterns == null)
            branchPatterns = storage.getBranchPatterns();
        // делаем все буквы маленькие и удаляем пробелы.
        String positionString = position.toLowerCase().replaceAll("\\s", "");
        // сопостовляем строчку с ключами словоря филиалов
        for (Map.Entry<String, Integer> entry : branchPatterns.entrySet()) {
            if (positionString.contains(entry.getKey()))
                return entry.getValue();
        }
        //Если не совпало ни с одним из паттернов значит это Администрация
        return 1;
    }


    void getReport(int month, int year) throws SQLException, IOException {

        LocalDate beginDate = LocalDate.of(year, month, 1);
        LocalDate endDate = LocalDate.of(year, month, LocalDate.of(year, month, 1).lengthOfMonth());

        XSSFWorkbook workbook = new XSSFWorkbook();
        ExcelStyles excelStyles = new ExcelStyles(workbook);

        List<String> branches = storage.getBranchesListInReport(beginDate, endDate, EventTypes.ОБУЧЕНИЕ);

        //Если за указанный периуд мероприяий не проводилось то завершение программы
        if (branches.isEmpty()){
            System.out.println("За период с " + beginDate + " по " + endDate + ". Мероприятий не проводилось.");
            System.exit(0);
        }

        for (String branch : branches) {
            XSSFSheet sheet = workbook.createSheet(branch);

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
            for (int i = 0; i < reportHeader.size(); i++) {
                switch (i) {
                    case 3:
                        if ("Администрация".equals(branch)) {
                            reportHeader.get(i).setValueStr("В Администрация OOO «Газпром трансгаз Москва»").addCell(sheet, excelStyles);
                        } else {
                            reportHeader.get(i).setValueStr(new StringBuilder().append("В филиал ").append(branch).toString()).addCell(sheet, excelStyles);
                        }
                        break;
                    case 5:
                        //todo исправить вывод месяца на вывод с маленькой буквы
                        reportHeader.get(i)
                                .setValueStr(new StringBuilder()
                                        .append("за ").append(beginDate.format(DateTimeFormatter.ofPattern("LLLL YYYY")))
                                        .append(" года").toString())
                                .addCell(sheet, excelStyles);
                        break;
                    default:
                        reportHeader.get(i).addCell(sheet, excelStyles);
                }
            }

            //-------Добавление основных данных в лист(людей)-------

            int rowNum = 10;
            List<String> events = storage.getEventsByBranchListInReport(branch, beginDate, endDate, EventTypes.ОБУЧЕНИЕ);

            for (String event : events) {
                int startRowMerged = rowNum;
                new CellData(rowNum,2, event,"t12alCCWB").addCell(sheet, excelStyles);

                Map<String, Integer> persons = storage.getPersonsInEventListInReport(branch, event, beginDate, endDate, EventTypes.ОБУЧЕНИЕ);
                for (Map.Entry<String, Integer> person : persons.entrySet()) {
                    new CellData(rowNum,0,rowNum-9,"t12alCCWB").addCell(sheet, excelStyles);
                    new CellData(rowNum,1, person.getKey(),"t12alLCWB").addCell(sheet, excelStyles);
                    new CellData(rowNum,3, person.getValue(),"t12alCCWB").addCell(sheet, excelStyles);
                    rowNum++;
                }
                if (persons.size() > 1) {
                    //todo выровнить ячейку по тексту
                    //обединение клеток с одним мероприятием
                    sheet.addMergedRegion(new CellRangeAddress(startRowMerged, rowNum - 1, 2, 2));
                }
            }

            //-------Добавление Итога и хвоста страницы-------

            // Обединение ячеек итога
            sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 2));

            List<CellData> reportFooter = storage.getTemplate("report_footer");
            for (int i = 0; i < reportFooter.size(); i++) {
                switch (i) {
                    case 3:
                        reportFooter.get(i)
                                .setValueStr("SUM(D11:D" + rowNum + ")")
                                .setIsFormula()
                                .addCell(sheet,excelStyles,rowNum);
                        break;
                    default:
                        reportFooter.get(i).addCell(sheet, excelStyles, rowNum);
                }
            }

            //Добавляем свойство "Разместить не более чем на 1 странице"
            XSSFPrintSetup printSetup = sheet.getPrintSetup();
            sheet.setFitToPage(true);
            printSetup.setFitHeight((short) 1);
            printSetup.setFitWidth((short) 1);
        }

        //Запись в файл
        try (FileOutputStream fileOut = new FileOutputStream("Отчет по обучению " + month + "_" + year + ".xlsx")) {
            workbook.write(fileOut);
        }
    }
}