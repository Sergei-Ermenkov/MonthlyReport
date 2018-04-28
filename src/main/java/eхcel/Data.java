package eхcel;

import data.DatePeriud;
import data.Event;
import data.EventTypes;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;
import data.Person;
import storage.SQLiteStorage;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.util.Map;

//TODO попробовать написать с помощю org.apache.poi.ss.util.CellUtil.createCell(Row row, int column, java.lang.String value)
public class Data {
    private final SQLiteStorage storage = new SQLiteStorage();
    private Map<String, String> branchPatterns;

    public void importExcel(String file) throws IOException, SQLException {
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
}