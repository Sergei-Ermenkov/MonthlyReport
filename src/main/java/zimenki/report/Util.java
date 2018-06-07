package zimenki.report;

import java.awt.font.FontRenderContext;
import java.awt.font.LineBreakMeasurer;
import java.awt.font.TextAttribute;
import java.text.AttributedString;

public class Util {
    public static int getHigh(String fontName, int fontSizeInPoints, String cellValue, float mergedCellWidthInPixels) {

        // Создайте объект Font с атрибутом Font (например, семейство шрифтов, размер шрифта и т. Д.) Для вычисления
        java.awt.Font currFont = new java.awt.Font(fontName, java.awt.Font.PLAIN, fontSizeInPoints);
        AttributedString attrStr = new AttributedString(cellValue);
        attrStr.addAttribute(TextAttribute.FONT, currFont);

        // Используйте LineBreakMeasurer для подсчета количества строк, необходимых для текста.
        FontRenderContext frc = new FontRenderContext(null, true, true);
        LineBreakMeasurer measurer = new LineBreakMeasurer(attrStr.getIterator(), frc);
        int nextPos;
        int lineCnt = 3;
        while (measurer.getPosition() < cellValue.length()) {
            nextPos = measurer.nextOffset(mergedCellWidthInPixels); // mergedCellWidthInPixels - максимальная ширина каждой строки
            lineCnt++;
            measurer.setPosition(nextPos);
        }
        return lineCnt;
        // Вышеупомянутое решение не обрабатывает символ новой строки, т. Е. «\n», и только
        // проверяется в горизонтальных сложенных ячейках.
    }

}
