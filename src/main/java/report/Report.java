package report;

import data.DatePeriud;

import eхcel.CellData;
import eхcel.ExcelStyles;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import storage.SQLiteStorage;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.*;

public abstract class Report {

    static final SQLiteStorage storage = new SQLiteStorage();

    final XSSFWorkbook workbook = new XSSFWorkbook();
    final ExcelStyles excelStyles = new ExcelStyles(workbook);
    final DatePeriud date;

    Report(DatePeriud date) {
        this.date = date;
    }

    abstract public void makeReport() throws SQLException, IOException;

    void checkCollection(Collection collection) {
        if (collection.isEmpty()) {
            throw new NullPointerException("За период с " + date.getBeginDate() + " по " + date.getEndDate() + ". Мероприятий не проводилось.");
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
        for (int i = 0; i < cellRangeAddresses.length; i++) {
            sheet.addMergedRegion(new CellRangeAddress(cellRangeAddresses[i][0],
                    cellRangeAddresses[i][1],
                    cellRangeAddresses[i][2],
                    cellRangeAddresses[i][3]));
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
        try (FileOutputStream fileOut = new FileOutputStream(name + date.getBeginDate().getMonthValue() + "_" + date.getBeginDate().getYear() + ".xlsx")) {
            workbook.write(fileOut);
        }
    }

}
