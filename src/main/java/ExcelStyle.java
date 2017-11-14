import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStyle {
    static {

    }

    private static XSSFFont getFont(XSSFWorkbook wb, int fontHeight, boolean underline, boolean bold){
        XSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short) fontHeight); //Устанавливает размер шрифта
        font.setFontName("Times New Roman"); //Задает имя шрифта
        if (underline){
            font.setUnderline(XSSFFont.U_SINGLE); //Подчеркивание текста линией
        }
        if (bold){
            font.setBold(true); //Задает жирный шрифт
        }
        return font;
    }

    private static void setBorder(XSSFCellStyle cellStyle, boolean fullBorder){
        cellStyle.setBorderBottom(BorderStyle.THIN); //рисует линию снизу
        if(fullBorder) {
            cellStyle.setBorderTop(BorderStyle.THIN); //рисует линию сверху
            cellStyle.setBorderLeft(BorderStyle.THIN); //рисует линию слева
            cellStyle.setBorderRight(BorderStyle.THIN); //рисует линию справа
        }
    }

    private static void setAlignment(XSSFCellStyle cellStyle, HorizontalAlignment HA, boolean wrapText){
        cellStyle.setAlignment(HA); //выравнивание шрифта в ячейке по горизонтали по центру
        if(wrapText) {
            cellStyle.setWrapText(true); //переноса текста в ячейке согласно ее размеру (ширине)
        }
    }

    private static void setAlignment(XSSFCellStyle cellStyle, HorizontalAlignment HA, VerticalAlignment VA){
        cellStyle.setAlignment(HA); //выравнивание шрифта в ячейке по горизонтали по центру
        cellStyle.setVerticalAlignment(VA); //выравнивание шрифта в ячейке по вертикали по центру
        cellStyle.setWrapText(true); //переноса текста в ячейке согласно ее размеру (ширине)
    }

    static XSSFCellStyle getT12Style(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, false));
        return style;
    }

    static XSSFCellStyle getT12unStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, true, false));
        return style;
    }

    static XSSFCellStyle getT9alRStyle(XSSFWorkbook wb){

        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,9, false, false));
        setAlignment(style, HorizontalAlignment.RIGHT, false);
        return style;
    }

    static XSSFCellStyle getT9alCStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,9, false, false));
        setAlignment(style, HorizontalAlignment.CENTER, false);
        return style;
    }

    static XSSFCellStyle getT9alLStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,9, false, false));
        setAlignment(style, HorizontalAlignment.LEFT, false);
        return style;
    }

    static XSSFCellStyle getT12alCStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, false));
        setAlignment(style, HorizontalAlignment.CENTER, false);
        return style;
    }

    static XSSFCellStyle getT12balLWStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, true));
        setAlignment(style, HorizontalAlignment.LEFT, true);
        return style;
    }

    static XSSFCellStyle getT12balCStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, true));
        setAlignment(style, HorizontalAlignment.CENTER, false);
        return style;
    }

    static XSSFCellStyle getT12balCWStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, true));
        setAlignment(style, HorizontalAlignment.CENTER, true);
        return style;
    }

    static XSSFCellStyle getT12alCWBStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, false));
        setAlignment(style, HorizontalAlignment.CENTER, true);
        setBorder(style, true);
        return style;
    }

    static XSSFCellStyle getT12balCWBStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, true));
        setAlignment(style, HorizontalAlignment.CENTER, true);
        setBorder(style, true);
        return style;
    }

    static XSSFCellStyle getT12balRBStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, true));
        setAlignment(style, HorizontalAlignment.RIGHT, false);
        setBorder(style, true);
        return style;
    }

    static XSSFCellStyle getT12alCBBStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, false));
        setAlignment(style, HorizontalAlignment.CENTER, false);
        setBorder(style, false);
        return style;
    }

    static XSSFCellStyle getNumStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb, 12, false, false));
        setAlignment(style, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(style, true);
        return style;
    }

    static XSSFCellStyle getFIOStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, false));
        setAlignment(style, HorizontalAlignment.LEFT, VerticalAlignment.CENTER);
        setBorder(style, true);
        return style;
    }

    static XSSFCellStyle getMeropStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb, 12, false, false));
        setAlignment(style, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(style, true);
        return style;
    }

    static XSSFCellStyle getTimeStyle(XSSFWorkbook wb){
        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(getFont(wb,12, false, false));
        setAlignment(style, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(style, true);
        return style;
    }
}
