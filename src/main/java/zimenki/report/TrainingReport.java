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
import java.util.Set;

public class TrainingReport extends Report {
    private static final int[][] columnWide = {{0, 1792}, {1, 9069}, {2, 7350}, {3, 4717}};
    private static final int[][] cellRangeAddresses = {{3,3,0,2},
            {5,5,0,3},
            {6,6,0,3}};
    private static final String header = "report_header";
    private static final int beginBodyRow = 10;
    private static final String footer = "footer";
    private static final String fileName = "Отчет по обучению_";

    public TrainingReport(LocalDate date){
        super(new DatePeriud(date));
    }

    @Override
    public void makeReport() throws SQLException, IOException {
        //Создаем Set всех филиалов (так как они могут повторяться) у которых было Обучение и Тех учеба
        // за указанный период

        Set<String> branches = storage.getBranchesStrings(report_periud, EventTypes.ОБУЧЕНИЕ);
        branches.addAll(storage.getBranchesStrings(report_periud, EventTypes.ТЕХУЧЕБА));

        checkCollection(branches, "обучения");

        for (String branch : branches) {
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(branch));

            //-------Форматирование листа и добавление шапки-------

            setColumnWideInExcel(sheet,columnWide);
            setMergedRegionInExcel(sheet, cellRangeAddresses);
            addTemplateToExcel(sheet, header);

            //Подстановка в шапку переменных полей
            addFilialNameToExcel(sheet, branch);
            addDateToExcel(sheet);

            //-------Добавление основных данных в лист(людей)-------

            int cursorRowNum = beginBodyRow;

            List<Event> events = storage.getEvents(report_periud, EventTypes.ОБУЧЕНИЕ, branch);
            events.addAll(storage.getEvents(report_periud, EventTypes.ТЕХУЧЕБА, branch));

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

        if (event.getNumberOfPersons() > 1) {
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
                .setCellValue("за " + report_periud.getMonthYearOfBeginDate() + " года");
    }

    private void addFilialNameToExcel(XSSFSheet sheet, String branch) {
        if ("Администрация".equals(branch)) {
            sheet.getRow(3).getCell(0).setCellValue("В Администрация OOO «Газпром трансгаз Москва»");
        } else {
            sheet.getRow(3).getCell(0).setCellValue("В филиал " + branch);
        }
    }

}
