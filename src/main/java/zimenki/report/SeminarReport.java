package zimenki.report;

import zimenki.data.DatePeriud;
import zimenki.data.Event;
import zimenki.data.EventTypes;
import zimenki.data.Person;
import zimenki.eхcel.CellData;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.sql.SQLException;
import java.time.LocalDate;
import java.util.List;

public class SeminarReport extends Report {
    private static final int[][] columnWide = {{0, 1792}, {1, 9069}, {2, 7350}, {3, 4717}};
    private static final int[][] cellRangeAddresses = {{3, 3, 0, 3},
            {4, 4, 0, 3},
            {5, 5, 0, 3}};
    private static final String header = "list_header";
    private static final int beginBodyRow = 9;
    private static final String footer = "footer";
    private static final String fileName = "Список участников семинаров_";

    public SeminarReport(LocalDate date) {
        super(new DatePeriud(date));
    }

    @Override
    public void makeReport() throws SQLException, IOException {

        List<Event> events = storage.getEvents(report_periud, EventTypes.СЕМИНАР);
        events.addAll(storage.getEvents(report_periud, EventTypes.ПРОЧЕЕ));

        checkCollection(events, "семинары");

        //изменение в переменных
        for (Event event : events) {
            XSSFSheet sheet = createSheet(WorkbookUtil.createSafeSheetName(event.getName()));

            //-------Форматирование листа и добавление шапки-------

            setColumnWideInExcel(sheet, columnWide);
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
            printSetup.setFitWidth((short) 1);
        }

        saveToFile(workbook, fileName);
    }

    private XSSFSheet createSheet(String sheetName) {
        return createSheet(new StringBuilder(sheetName), 1);
    }

    private XSSFSheet createSheet(StringBuilder sheetName, int i) {
        if (workbook.getSheet(sheetName.toString()) == null)
            return workbook.createSheet(sheetName.toString());
        else {
            String iS = Integer.toString(i);
            sheetName.replace(sheetName.length()-iS.length(), sheetName.length(), iS);
            return createSheet(sheetName, i+1);
        }
    }

    private int addPersonsToExcel(XSSFSheet sheet, Event event, int cursorRowNum) {
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
                .setCellValue("в период " + event.getDate().intersection(report_periud));
    }

    private void addNameSeminarToExcel(XSSFSheet sheet, Event event) {
        sheet.getRow(4).getCell(0)
                .setCellValue("участников мероприятия: " + event.getName());
    }

    private void setHighNameSeminarRow(XSSFSheet sheet, Event event) {
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
