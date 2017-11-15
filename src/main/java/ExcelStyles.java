import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class ExcelStyles {
    XSSFCellStyle T12Style;
    XSSFCellStyle T12unStyle;
    XSSFCellStyle T9alRStyle;
    XSSFCellStyle T9alCStyle;
    XSSFCellStyle T9alLStyle;
    XSSFCellStyle T12alCStyle;
    XSSFCellStyle T12balLWStyle;
    XSSFCellStyle T12balCStyle;
    XSSFCellStyle T12balCWStyle;
    XSSFCellStyle T12alCWBStyle;
    XSSFCellStyle T12balCWBStyle;
    XSSFCellStyle T12balRBStyle;
    XSSFCellStyle T12alCBBStyle;
    XSSFCellStyle NumStyle;
    XSSFCellStyle FIOStyle;
    XSSFCellStyle MeropStyle;
    XSSFCellStyle TimeStyle;

    ExcelStyles(XSSFWorkbook wb) {
        T12Style = wb.createCellStyle();
        T12Style.setFont(getFont(wb, 12, false, false));

        T12unStyle = wb.createCellStyle();
        T12unStyle.setFont(getFont(wb, 12, true, false));

        T9alRStyle = wb.createCellStyle();
        T9alRStyle.setFont(getFont(wb, 9, false, false));
        setAlignment(T9alRStyle, HorizontalAlignment.RIGHT, false);

        T9alCStyle = wb.createCellStyle();
        T9alCStyle.setFont(getFont(wb, 9, false, false));
        setAlignment(T9alCStyle, HorizontalAlignment.CENTER, false);

        T9alLStyle = wb.createCellStyle();
        T9alLStyle.setFont(getFont(wb, 9, false, false));
        setAlignment(T9alLStyle, HorizontalAlignment.LEFT, false);

        T12alCStyle = wb.createCellStyle();
        T12alCStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(T12alCStyle, HorizontalAlignment.CENTER, false);

        T12balLWStyle = wb.createCellStyle();
        T12balLWStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(T12balLWStyle, HorizontalAlignment.LEFT, true);

        T12balCStyle = wb.createCellStyle();
        T12balCStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(T12balCStyle, HorizontalAlignment.CENTER, false);

        T12balCWStyle = wb.createCellStyle();
        T12balCWStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(T12balCWStyle, HorizontalAlignment.CENTER, true);

        T12alCWBStyle = wb.createCellStyle();
        T12alCWBStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(T12alCWBStyle, HorizontalAlignment.CENTER, true);
        setBorder(T12alCWBStyle, true);

        T12balCWBStyle = wb.createCellStyle();
        T12balCWBStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(T12balCWBStyle, HorizontalAlignment.CENTER, true);
        setBorder(T12balCWBStyle, true);

        T12balRBStyle = wb.createCellStyle();
        T12balRBStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(T12balRBStyle, HorizontalAlignment.RIGHT, false);
        setBorder(T12balRBStyle, true);

        T12alCBBStyle = wb.createCellStyle();
        T12alCBBStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(T12alCBBStyle, HorizontalAlignment.CENTER, false);
        setBorder(T12alCBBStyle, false);

        NumStyle = wb.createCellStyle();
        NumStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(NumStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(NumStyle, true);

        FIOStyle = wb.createCellStyle();
        FIOStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(FIOStyle, HorizontalAlignment.LEFT, VerticalAlignment.CENTER);
        setBorder(FIOStyle, true);

        MeropStyle = wb.createCellStyle();
        MeropStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(MeropStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(MeropStyle, true);

        TimeStyle = wb.createCellStyle();
        TimeStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(TimeStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(TimeStyle, true);
}

    private XSSFFont getFont(XSSFWorkbook wb, int fontHeight, boolean underline, boolean bold) {
        XSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short) fontHeight); //Устанавливает размер шрифта
        font.setFontName("Times New Roman"); //Задает имя шрифта
        if (underline) {
            font.setUnderline(XSSFFont.U_SINGLE); //Подчеркивание текста линией
        }
        if (bold) {
            font.setBold(true); //Задает жирный шрифт
        }
        return font;
    }

    private void setBorder(XSSFCellStyle cellStyle, boolean fullBorder) {
        cellStyle.setBorderBottom(BorderStyle.THIN); //рисует линию снизу
        if (fullBorder) {
            cellStyle.setBorderTop(BorderStyle.THIN); //рисует линию сверху
            cellStyle.setBorderLeft(BorderStyle.THIN); //рисует линию слева
            cellStyle.setBorderRight(BorderStyle.THIN); //рисует линию справа
        }
    }

    private void setAlignment(XSSFCellStyle cellStyle, HorizontalAlignment HA, boolean wrapText) {
        cellStyle.setAlignment(HA); //выравнивание шрифта в ячейке по горизонтали по центру
        if (wrapText) {
            cellStyle.setWrapText(true); //переноса текста в ячейке согласно ее размеру (ширине)
        }
    }

    private void setAlignment(XSSFCellStyle cellStyle, HorizontalAlignment HA, VerticalAlignment VA) {
        cellStyle.setAlignment(HA); //выравнивание шрифта в ячейке по горизонтали по центру
        cellStyle.setVerticalAlignment(VA); //выравнивание шрифта в ячейке по вертикали по центру
        cellStyle.setWrapText(true); //переноса текста в ячейке согласно ее размеру (ширине)
    }
}
