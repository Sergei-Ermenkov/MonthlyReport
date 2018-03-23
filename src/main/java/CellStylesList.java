import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
Стили для форматирования ячеек
t12 - Times New Roman, шрит 12
un - линия снизу
alR


 */
class CellStylesList {
    XSSFCellStyle t12Style;
    XSSFCellStyle t12unStyle;
    XSSFCellStyle t9alRStyle;
    XSSFCellStyle t9alCStyle;
    XSSFCellStyle t9alLStyle;
    XSSFCellStyle t12alCStyle;
    XSSFCellStyle t12balLWStyle;
    XSSFCellStyle t12balCStyle;
    XSSFCellStyle t12balCWStyle;
    XSSFCellStyle t12alCWBStyle;
    XSSFCellStyle t12balCWBStyle;
    XSSFCellStyle t12balRBStyle;
    XSSFCellStyle t12alCBBStyle;
    XSSFCellStyle numStyle;
    XSSFCellStyle fioStyle;
    XSSFCellStyle topicStyle;
    XSSFCellStyle timeStyle;

    CellStylesList(XSSFWorkbook wb) {
        t12Style = wb.createCellStyle();
        t12Style.setFont(getFont(wb, 12, false, false));

        t12unStyle = wb.createCellStyle();
        t12unStyle.setFont(getFont(wb, 12, true, false));

        t9alRStyle = wb.createCellStyle();
        t9alRStyle.setFont(getFont(wb, 9, false, false));
        setAlignment(t9alRStyle, HorizontalAlignment.RIGHT, false);

        t9alCStyle = wb.createCellStyle();
        t9alCStyle.setFont(getFont(wb, 9, false, false));
        setAlignment(t9alCStyle, HorizontalAlignment.CENTER, false);

        t9alLStyle = wb.createCellStyle();
        t9alLStyle.setFont(getFont(wb, 9, false, false));
        setAlignment(t9alLStyle, HorizontalAlignment.LEFT, false);

        t12alCStyle = wb.createCellStyle();
        t12alCStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(t12alCStyle, HorizontalAlignment.CENTER, false);

        t12balLWStyle = wb.createCellStyle();
        t12balLWStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(t12balLWStyle, HorizontalAlignment.LEFT, true);

        t12balCStyle = wb.createCellStyle();
        t12balCStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(t12balCStyle, HorizontalAlignment.CENTER, false);

        t12balCWStyle = wb.createCellStyle();
        t12balCWStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(t12balCWStyle, HorizontalAlignment.CENTER, true);

        t12alCWBStyle = wb.createCellStyle();
        t12alCWBStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(t12alCWBStyle, HorizontalAlignment.CENTER, true);
        setBorder(t12alCWBStyle, true);

        t12balCWBStyle = wb.createCellStyle();
        t12balCWBStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(t12balCWBStyle, HorizontalAlignment.CENTER, true);
        setBorder(t12balCWBStyle, true);

        t12balRBStyle = wb.createCellStyle();
        t12balRBStyle.setFont(getFont(wb, 12, false, true));
        setAlignment(t12balRBStyle, HorizontalAlignment.RIGHT, false);
        setBorder(t12balRBStyle, true);

        t12alCBBStyle = wb.createCellStyle();
        t12alCBBStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(t12alCBBStyle, HorizontalAlignment.CENTER, false);
        setBorder(t12alCBBStyle, false);

        numStyle = wb.createCellStyle();
        numStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(numStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(numStyle, true);

        fioStyle = wb.createCellStyle();
        fioStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(fioStyle, HorizontalAlignment.LEFT, VerticalAlignment.CENTER);
        setBorder(fioStyle, true);

        topicStyle = wb.createCellStyle();
        topicStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(topicStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(topicStyle, true);

        timeStyle = wb.createCellStyle();
        timeStyle.setFont(getFont(wb, 12, false, false));
        setAlignment(timeStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(timeStyle, true);
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
