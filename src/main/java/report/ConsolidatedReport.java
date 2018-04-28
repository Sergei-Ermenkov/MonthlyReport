package report;

import data.DatePeriud;
import data.Event;
import data.EventTypes;
import eхcel.CellData;
import org.apache.poi.ss.usermodel.PrintOrientation;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.sql.SQLException;
import java.util.Iterator;
import java.util.List;

public class ConsolidatedReport extends Report {
    private static final int[][] columnWide = {{0, 1865}, {1, 12397}, {2, 3766}, {3, 4169}, {4, 9947}, {5, 3803}, {6, 3803}, {7, 4827}, {8, 6144}};
    private static final short[][] rowHigh = {{0, 330}, {1, 195}, {5, 420}, {6, 150}, {9, 1575}};
    private static final int[][] cellRangeAddresses = {{0, 0, 7, 8},
            {3, 3, 6, 8},
            {4, 4, 6, 8},
            {5, 5, 7, 8},
            {7, 7, 0, 8}};
    private static final String header = "consolidated_header";
    private static final int beginBodyRow = 11;
    private static final String footer = "consolidated_footer";
    private static final String fileName = "Сводный отчет УЧ Зименки_";

    public ConsolidatedReport(int month, int year) {
        super(new DatePeriud(month, year));
    }

    @Override
    public void makeReport() throws SQLException, IOException {

        List<Event> events = storage.getEvents(date);

        checkCollection(events);

        combineTech(events);

        //todo: проверить данные
        XSSFSheet sheet = workbook.createSheet("Сводный отчет");
        sheet.getPrintSetup().setOrientation(PrintOrientation.LANDSCAPE);

        //----> Шапка <-----
        setColumnWideInExcel(sheet, columnWide);
        setRowHighInExcel(sheet, rowHigh);

        setMergedRegionInExcel(sheet, cellRangeAddresses);

        addTemplateToExcel(sheet, header);


        sheet.getRow(7).getCell(0)
                .setCellValue("Сводный отчет о проведенных мероприятиях в УЧ «Зименки» за " + date.getMonthYearOfBeginDate() + " года");
        //------------------

        //----> Тело <------
        int cursorRowNum = addEventToExcel(sheet, events, beginBodyRow);
        //------------------

        //----> Хвост <-----

        // Обединение ячеек итога
        //TODO: вынести в родителя
        //строка всего
        sheet.addMergedRegion(new CellRangeAddress(cursorRowNum, cursorRowNum, 0, 1));
        //строки фамилий вподписи

        sheet.addMergedRegion(new CellRangeAddress(cursorRowNum + 2, cursorRowNum + 2, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(cursorRowNum + 3, cursorRowNum + 3, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(cursorRowNum + 6, cursorRowNum + 6, 3, 4));
        sheet.addMergedRegion(new CellRangeAddress(cursorRowNum + 7, cursorRowNum + 7, 3, 4));

        // Добавление хвоста в лист

        addTemplateToExcel(sheet, footer, cursorRowNum);

        //Подстановка в формулы в итог
        sheet.getRow(cursorRowNum).getCell(5).setCellFormula("SUM(F12:F" + cursorRowNum + ")");
        sheet.getRow(cursorRowNum).getCell(6).setCellFormula("SUM(G12:G" + cursorRowNum + ")");
        sheet.getRow(cursorRowNum).getCell(7).setCellFormula("SUM(H12:H" + cursorRowNum + ")");
        sheet.getRow(cursorRowNum).getCell(8).setCellFormula("SUM(I12:I" + cursorRowNum + ")");

        sheet.getRow(cursorRowNum + 3).setHeight((short) 220);
        sheet.getRow(cursorRowNum + 4).setHeight((short) 210);
        sheet.getRow(cursorRowNum + 6).setHeight((short) 390);
        sheet.getRow(cursorRowNum + 7).setHeight((short) 220);

        //Добавляем свойство "Разместить не более чем на 1 странице"
        XSSFPrintSetup printSetup = sheet.getPrintSetup();
        sheet.setFitToPage(true);
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);

        saveToFile(workbook, fileName);
    }

    private void combineTech(List<Event> events){
        Event tech = null;
        Iterator<Event> eventIterator = events.iterator();
        while (eventIterator.hasNext()){
            Event event = eventIterator.next();
            if (event.getType() == EventTypes.ТЕХУЧЕБА){
                if(tech == null)
                    tech = new Event(event.getType(), "Проведение технической учебы и производственных инструктажей в Учебной части (Зименки) Учебно-производственного центра.",
                            event.getDecree(), event.getDate());
                tech.addPersons(event.getPersons());
                eventIterator.remove();
            }
        }
        events.add(tech);
    }

    private int addEventToExcel(XSSFSheet sheet, List<Event> events, int cursorRowNum){

        for (Event event : events) {
            new CellData(cursorRowNum, 0, cursorRowNum - 10, "t12alCCWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 1, event.getName(), "t12alLHWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 2, event.getDate().toString(), "t12alCHWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 3, event.getNumberOfDayFrom(date), "t12alCCWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 4, event.getDecree(), "t12alLHWB").addCell(sheet, excelStyles);
            new CellData(cursorRowNum, 5, event.getNumberOfPersons(), "t12alCCWB").addCell(sheet, excelStyles);
            switch (event.getType()){
                case ТЕХУЧЕБА:
                case ОБУЧЕНИЕ:
                    new CellData(cursorRowNum, 6, event.getSumManHours(), "t12alCCWB").addCell(sheet, excelStyles);
                    sheet.getRow(cursorRowNum).createCell(7).setCellStyle(excelStyles.getStyle("t12alCCWB"));
                    sheet.getRow(cursorRowNum).createCell(8).setCellStyle(excelStyles.getStyle("t12alCCWB"));
                    break;
                case СЕМИНАР:
                    sheet.getRow(cursorRowNum).createCell(6).setCellStyle(excelStyles.getStyle("t12alCCWB"));
                    new CellData(cursorRowNum, 7, event.getSumManHours(), "t12alCCWB").addCell(sheet, excelStyles);
                    sheet.getRow(cursorRowNum).createCell(8).setCellStyle(excelStyles.getStyle("t12alCCWB"));
                    break;
                case ПРОЧЕЕ:
                    sheet.getRow(cursorRowNum).createCell(6).setCellStyle(excelStyles.getStyle("t12alCCWB"));
                    sheet.getRow(cursorRowNum).createCell(7).setCellStyle(excelStyles.getStyle("t12alCCWB"));
                    new CellData(cursorRowNum, 8, event.getSumManHours(), "t12alCCWB").addCell(sheet, excelStyles);
            }
            cursorRowNum++;
        }
     return cursorRowNum;
    }


}
