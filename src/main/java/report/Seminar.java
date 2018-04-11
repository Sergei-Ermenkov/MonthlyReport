package report;

import data.DatePeriud;
import data.Event;
import data.EventTypes;
import data.Person;
import eхcel.CellData;
import eхcel.Util;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.sql.SQLException;
import java.util.List;

public class Seminar extends Report {
    final int[][] cellRangeAddresses = {{3, 3, 0, 3},
            {4, 4, 0, 3},
            {5, 5, 0, 3}};
    final String header = "list_header";
    final int beginBodyRow = 9;
    final String footer = "footer";

    final String fileName = "Список участников семинаров_";

    public Seminar(int month, int year) {
        super(new DatePeriud(month, year));
    }

    @Override
    public void makeReport() throws SQLException, IOException {

        List<Event> events = storage.getEvents(date, EventTypes.СЕМИНАР);
        events.addAll(storage.getEvents(date, EventTypes.ПРОЧЕЕ));

        checkCollection(events);

        //изменение в переменных
        //todo добавить проверку на совпадение нового имени с именем в книги
        for (Event event : events) {
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(event.getName()));

            //-------Форматирование листа и добавление шапки-------

            setWideColumnInExcel(sheet);
            setMergedRegionInExcel(sheet, cellRangeAddresses);
            addTemplateToExcel(sheet, header);


            //Подстановка в шапку переменных полей
            addNameSeminarToExcel(sheet, event);
            setHighNameSeminarRow(sheet, event);
            addDatePeriudToExcel(sheet, event);


            //-------Добавление основных данных в лист(людей)-------

            int cursorRowNum = addPersonsToExcel(sheet, event, beginBodyRow);

            //-------Добавление Итога и хвоста страницы-------

            // Обединение ячеек итога
            sheet.addMergedRegion(new CellRangeAddress(cursorRowNum, cursorRowNum, 0, 2));
            // Добавление хвоста в лист
            addTemplateToExcel(sheet, footer, cursorRowNum);

            //Подстановка в формулы в итог
            sheet.getRow(cursorRowNum).getCell(3).setCellFormula("SUM(D10:D" + cursorRowNum + ")");

            //Добавляем свойство "Разместить не более чем на 1 странице"
            XSSFPrintSetup printSetup = sheet.getPrintSetup();
            sheet.setFitToPage(true);
            //todo убран ограничитель на 1 страницу по высоте
            printSetup.setFitWidth((short) 1);
        }

        saveToFile(workbook, fileName);
    }

    int addPersonsToExcel(XSSFSheet sheet, Event event, int cursorRowNum) throws SQLException {
        for (Person person : event.getPersons()) {
            new CellData(cursorRowNum, 0, cursorRowNum - 8, "t12alCCWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 1, person.getName(), "t12alLCWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 2, person.getPosition(), "t12alCCWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 3, person.getManHours(), "t12alCCWB").addCell(sheet, excelStyles);
            cursorRowNum++;
        }
        return cursorRowNum;
    }

    private void addDatePeriudToExcel(XSSFSheet sheet, Event event) {
        sheet.getRow(5).getCell(0)
                .setCellValue("в период " + event.getDate());
    }

    private void addNameSeminarToExcel(XSSFSheet sheet, Event event) {
        sheet.getRow(4).getCell(0)
                .setCellValue("участников мероприятия: " + event.getName());
    }

    private void setHighNameSeminarRow(XSSFSheet sheet, Event event) {
        //todo заменить ручное указание стиля
        XSSFCellStyle style = excelStyles.getStyle("t12balCW");
        int rowsHigh = Util.getHigh(style.getFont().getFontName(),
                style.getFont().getFontHeightInPoints(),
                event.getName(),
                sheet.getColumnWidthInPixels(0) +
                        sheet.getColumnWidthInPixels(1) +
                        sheet.getColumnWidthInPixels(2) +
                        sheet.getColumnWidthInPixels(3));
        if (rowsHigh > 1)
            sheet.getRow(4).setHeightInPoints(sheet.getDefaultRowHeightInPoints() * rowsHigh);
    }
}
