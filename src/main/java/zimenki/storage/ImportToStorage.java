package zimenki.storage;

import zimenki.data.DatePeriud;
import zimenki.data.Event;
import zimenki.data.EventTypes;
import zimenki.data.Person;
import zimenki.eхcel.NotIntersectDateException;
import zimenki.eхcel.WrongImportFilenameException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.Map;

public class ImportToStorage {
    private final SQLiteStorage storage = new SQLiteStorage();
    private Map<String, String> branchPatterns;

    public void importFromExcel(File file) throws IOException, SQLException {
        DatePeriud reportDate;
        try {
            reportDate = getDateFromFileName(file);
        } catch (NumberFormatException | ArrayIndexOutOfBoundsException | java.time.DateTimeException e){
            throw new WrongImportFilenameException("Некорректное имя импортируемого файла.");
        }

        try (FileInputStream fileInputStream = new FileInputStream(file);
             XSSFWorkbook excelBook = new XSSFWorkbook(fileInputStream)) {
            validImportData(excelBook);

            for (Sheet sheet : excelBook) {
                Event event = getEventFromExcel(sheet);
                if (!event.getDate().isIntersection(reportDate))
                    throw new NotIntersectDateException("Дата отчетного периуда не пересекается с датой мероприятия.");
                for (int i = 2; i < sheet.getLastRowNum() + 1; i++) {
                    Row row = sheet.getRow(i);
                    event.addPerson(getPersonFromExcel(row));
                }
                storage.addEventPersons(event, reportDate);
            }
        }
    }

    private DatePeriud getDateFromFileName(File file){
        String[] parse = file.getName().split("[_.]");
        int year = Integer.parseInt(parse[2]);
        if ((year < 2000) | (year > 2100)){
            throw new java.time.DateTimeException("Неверно указан год в названии файла");
        }
        return new DatePeriud(Integer.parseInt(parse[1]), year);
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
}