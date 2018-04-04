package eхcel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class CellData {
    private int rowNum;
    private int cellNum;
    private String valueStr;
    private int valueInt;
    private String style;
    private int typeData;

    public CellData(int rowNum, int cellNum, String valueStr, String style) {
        this(rowNum, cellNum, valueStr, style, false);
    }

    public CellData(int rowNum, int cellNum, String valueStr, String style, boolean isFormula) {
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
        XSSFCell cell = row.createCell(cellNum);
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

    public CellData setValueStr(String valueStr) {
        this.valueStr = valueStr;
        return this;
    }

    public CellData setIsFormula() {
        this.typeData = 3;
        return this;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("CellData{ rowNum=").append(rowNum).append(", cellNum=").append(cellNum).append(", valueStr='")
                .append(valueStr).append(", style='").append(style).append("'}");
        return sb.toString();
    }
}
