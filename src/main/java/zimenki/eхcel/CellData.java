package zimenki.eхcel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class CellData {
    private final int rowNum;
    private final int cellNum;
    private String valueStr;
    private int valueInt;
    private final String style;
    private int typeData;

    public CellData(int rowNum, int cellNum, String valueStr, String style) {
        this(rowNum, cellNum, valueStr, style, false);
    }

    private CellData(int rowNum, int cellNum, String valueStr, String style, boolean isFormula) {
        this.rowNum = rowNum;
        this.cellNum = cellNum;
        this.valueStr = valueStr;
        this.style = style;
        typeData = (isFormula) ? 3 : 1;
    }

    public CellData(int rowNum, int cellNum, int valueInt, String style) {
        this.rowNum = rowNum;
        this.cellNum = cellNum;
        this.valueInt = valueInt;
        this.style = style;
        typeData = 2;
    }

    public void addCell(XSSFSheet sheet, ExcelStyles excelStyles) {
        addCell(sheet, excelStyles, 0);
    }

    public void addCell(XSSFSheet sheet, ExcelStyles excelStyles, int shift) {
        XSSFRow row = sheet.getRow(rowNum + shift);
        if (row == null)
            row = sheet.createRow(rowNum + shift);

        XSSFCell cell = row.getCell(cellNum);
        if (cell == null)
            cell = row.createCell(cellNum);

        switch (typeData) {
            case 1:
                cell.setCellValue(valueStr);
                break;
            case 2:
                cell.setCellValue(valueInt);
                break;
            case 3:
                cell.setCellFormula(valueStr);
        }
        cell.setCellStyle(excelStyles.getStyle(style));
    }
}
