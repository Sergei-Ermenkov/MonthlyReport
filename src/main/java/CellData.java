import org.apache.poi.xssf.usermodel.XSSFCellStyle;

class CellData {
    private int numColumn;
    private String value;
    private XSSFCellStyle style;

    CellData(int numColumn, String value, XSSFCellStyle style) {
        this.numColumn = numColumn;
        this.value = value;
        this.style = style;
    }

    int getNumColumn() {
        return numColumn;
    }

    String getValue() {
        return value;
    }

    XSSFCellStyle getStyle() {
        return style;
    }

    void setValue(String value) {
        this.value = value;
    }
}
