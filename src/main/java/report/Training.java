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
import java.util.Set;

public class Training extends Report {
    final int[][] cellRangeAddresses = {{3,3,0,2},
            {5,5,0,3},
            {6,6,0,3}};

    final String header = "report_header";
    final int beginBodyRow = 10;
    final String footer = "footer";

    final String fileName = "Отчет по обучению_";

    public Training(int month, int year){
        super(new DatePeriud(month, year));
    }

    @Override
    public void makeReport() throws SQLException, IOException {
        //Создаем Set всех филиалов (так как они могут повторяться) у которых было Обучение и Тех учеба
        // за указанный период
        Set<String> branches = storage.getBranchesStrings(date, EventTypes.ОБУЧЕНИЕ);
        branches.addAll(storage.getBranchesStrings(date, EventTypes.ТЕХУЧЕБА));

        checkCollection(branches);

        for (String branch : branches) {
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(branch));

            //-------Форматирование листа и добавление шапки-------

            setWideColumnInExcel(sheet);
            setMergedRegionInExcel(sheet, cellRangeAddresses);
            addTemplateToExcel(sheet, header);

            //Подстановка в шапку переменных полей
            addFilialNameToExcel(sheet, branch);
            addDateToExcel(sheet);

            //-------Добавление основных данных в лист(людей)-------

            int cursorRowNum = beginBodyRow;

            List<Event> events = storage.getEvents(date, EventTypes.ОБУЧЕНИЕ, branch);
            events.addAll(storage.getEvents(date, EventTypes.ТЕХУЧЕБА, branch));

            for (Event event : events) {
                cursorRowNum = addPersonsToExcel(sheet, event, cursorRowNum);
            }

            //-------Добавление Итога и хвоста страницы-------

            // Обединение ячеек итога
            sheet.addMergedRegion(new CellRangeAddress(cursorRowNum, cursorRowNum, 0, 2));
            // Добавление хвоста в лист
            addTemplateToExcel(sheet, footer, cursorRowNum);

            //Подстановка в формулы в итог
            sheet.getRow(cursorRowNum).getCell(3).setCellFormula("SUM(D11:D" + cursorRowNum + ")");

            //Добавляем свойство "Разместить не более чем на 1 странице"
            XSSFPrintSetup printSetup = sheet.getPrintSetup();
            sheet.setFitToPage(true);
            printSetup.setFitHeight((short) 1);
            printSetup.setFitWidth((short) 1);
        }

        saveToFile(workbook, fileName);
    }

    private int addPersonsToExcel(XSSFSheet sheet, Event event, int cursorRowNum) {
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
            setHighNameTrainingRow(sheet, event, cursorRowNum, startRowMerged);
            //обединение клеток с одним мероприятием
            sheet.addMergedRegion(new CellRangeAddress(startRowMerged, cursorRowNum - 1, 2, 2));
        }
        return cursorRowNum;
    }

    private void setHighNameTrainingRow(XSSFSheet sheet, Event event, int cursorRowNum, int startRowMerged) {
        XSSFCellStyle style = excelStyles.getStyle("t12alCCWB");
        int rowsHigh = Util.getHigh(style.getFont().getFontName(),
                style.getFont().getFontHeightInPoints(),
                event.getName(),
                sheet.getColumnWidthInPixels(2));
        int rowsIs = cursorRowNum - startRowMerged;
        float newHigh = (sheet.getDefaultRowHeightInPoints() * rowsHigh) / rowsIs;
        if (rowsHigh > rowsIs) {
            for (int i = startRowMerged; i < cursorRowNum; i++) {
                sheet.getRow(i).setHeightInPoints(newHigh);
            }
        }
    }

    private void addDateToExcel(XSSFSheet sheet) {

        sheet.getRow(6).getCell(0)
                .setCellValue("за " + date.getMonthYearOfBeginDate() + " года");
    }

    private void addFilialNameToExcel(XSSFSheet sheet, String branch) {
        if ("Администрация".equals(branch)) {
            sheet.getRow(3).getCell(0).setCellValue("В Администрация OOO «Газпром трансгаз Москва»");
        } else {
            sheet.getRow(3).getCell(0).setCellValue("В филиал " + branch);
        }
    }

}
