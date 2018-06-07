package zimenki.report;

import zimenki.data.DatePeriud;
import zimenki.eхcel.CellData;
import zimenki.eхcel.ExcelStyles;
import zimenki.storage.SQLiteStorage;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.*;

public abstract class Report {

    static final SQLiteStorage storage = new SQLiteStorage();

    final XSSFWorkbook workbook = new XSSFWorkbook();
    final ExcelStyles excelStyles = new ExcelStyles(workbook);
    final DatePeriud report_periud;

    Report(DatePeriud report_periud) {
        this.report_periud = report_periud;
    }

    abstract public void makeReport() throws SQLException, IOException;

    void checkCollection(Collection collection, String typeReport) {
        if (collection.isEmpty()) {
            throw new NullEventsException("За период с " + report_periud.getBeginDate() + " по " + report_periud.getEndDate() + ".\n Отсутстуют " + typeReport);
        }

    }

    void setColumnWideInExcel(XSSFSheet sheet, int[][] columnWide) {
        for (int[] wide : columnWide) {
            sheet.setColumnWidth(wide[0], wide[1]);
        }
    }

    void setRowHighInExcel(XSSFSheet sheet, short[][] rowHigh) {
        for (short[] high : rowHigh) {
            XSSFRow row = sheet.getRow(high[0]);
            if (row == null)
                row = sheet.createRow(high[0]);
            row.setHeight(high[1]);
        }
    }

    void setMergedRegionInExcel(XSSFSheet sheet, int[][] cellRangeAddresses) {
        for (int[] cellRangeAddress : cellRangeAddresses) {
            sheet.addMergedRegion(new CellRangeAddress(cellRangeAddress[0],
                    cellRangeAddress[1],
                    cellRangeAddress[2],
                    cellRangeAddress[3]));
        }
    }

    void addTemplateToExcel(XSSFSheet sheet, String template) throws SQLException{
        addTemplateToExcel(sheet, template,0);
    }

    void addTemplateToExcel(XSSFSheet sheet, String template, int cursorRowNum) throws SQLException {
        List<CellData> listHeader = storage.getTemplate(template);
        for (CellData cellData : listHeader) {
            cellData.addCell(sheet, excelStyles, cursorRowNum);
        }
    }

    void saveToFile(Workbook workbook, String name) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(name + report_periud.getBeginDate().getMonthValue() + "_" + report_periud.getBeginDate().getYear() + ".xlsx")) {
            workbook.write(fileOut);
        }
    }



}
