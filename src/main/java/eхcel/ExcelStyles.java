package eхcel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

/*
Стили для форматирования ячеек
Пример:
t12balCWB

t12........ - Times New Roman 12
.t9........ - Times New Roman 9
...u....... - подчеркивание
...b....... - выделение жирным
....alC.... - выравнивание по горизонтали по центру
....alR.... - выравнивание по горизонтали по правому краю
....alL.... - выравнивание по горизонтали по левому краю
....alCC... - выравнивание по горизонтали по центру, по вертикали по центру
....alLC... - выравнивание по горизонтали по левому краю, по вертикали по центру
....alLH... - выравнивание по горизонтали по левому краю, по вертикали по верхнему краю
........W.. - переноса текста в ячейке согласно ее размеру (ширине)
.........B. - Обрамление ячейки линиями со всех сторон
.........BB - Обрамление ячейки линиями снизу
 */
public class ExcelStyles {
    private final Map<String, XSSFCellStyle> cellStylesMap = new HashMap<>();
    public ExcelStyles(XSSFWorkbook wb) {

        //-------Шрифты-------

        XSSFFont t9F = wb.createFont();
        t9F.setFontHeightInPoints((short) 9);   //Устанавливает размер шрифта
        t9F.setFontName("Times New Roman");     //Задает имя шрифта

        XSSFFont t12F = wb.createFont();
        t12F.setFontHeightInPoints((short) 12); //Устанавливает размер шрифта
        t12F.setFontName("Times New Roman");    //Задает имя шрифта

        XSSFFont t12uF = wb.createFont();
        t12uF.setFontHeightInPoints((short) 12); //Устанавливает размер шрифта
        t12uF.setFontName("Times New Roman");    //Задает имя шрифта
        t12uF.setUnderline(XSSFFont.U_SINGLE);   //Подчеркивание текста линией

        XSSFFont t12bF = wb.createFont();
        t12bF.setFontHeightInPoints((short) 12); //Устанавливает размер шрифта
        t12bF.setFontName("Times New Roman");    //Задает имя шрифта
        t12bF.setBold(true);                     //Задает жирный шрифт


        //-------Стили шаблоны-------

        XSSFCellStyle b = wb.createCellStyle();
        b.setBorderBottom(BorderStyle.THIN);         //Обрамление линиями снизу
        b.setBorderTop(BorderStyle.THIN);            //Обрамление линиями сверху
        b.setBorderLeft(BorderStyle.THIN);           //Обрамление линиями слева
        b.setBorderRight(BorderStyle.THIN);          //Обрамление линиями справа

        //-------Стили-------

        //t9alC
        XSSFCellStyle t9alC = wb.createCellStyle();
        t9alC.setFont(t9F);
        t9alC.setAlignment(HorizontalAlignment.CENTER);     //выравнивание по горизонтали по центру
        cellStylesMap.put("t9alC", t9alC);

        //t9alR
        XSSFCellStyle t9alR = wb.createCellStyle();
        t9alR.setFont(t9F);
        t9alR.setAlignment(HorizontalAlignment.RIGHT);      //выравнивание по горизонтали по праваму краю
        cellStylesMap.put("t9alR",t9alR);

        //t9alL
        XSSFCellStyle t9alL = wb.createCellStyle();
        t9alL.setFont(t9F);
        t9alL.setAlignment(HorizontalAlignment.LEFT);       //выравнивание по горизонтали по левому краю
        cellStylesMap.put("t9alL",t9alL);

        //t12
        XSSFCellStyle t12 = wb.createCellStyle();
        t12.setFont(t12F);
        cellStylesMap.put("t12",t12);

        //t12u
        XSSFCellStyle t12u = wb.createCellStyle();
        t12u.setFont(t12uF);
        cellStylesMap.put("t12u",t12u);

        //t12ualC
        XSSFCellStyle t12ualC = wb.createCellStyle();
        t12ualC.setFont(t12uF);
        t12ualC.setAlignment(HorizontalAlignment.CENTER);    //выравнивание по горизонтали по центру
        cellStylesMap.put("t12ualC",t12ualC);

        //t12alC
        XSSFCellStyle t12alC = wb.createCellStyle();
        t12alC.setFont(t12F);
        t12alC.setAlignment(HorizontalAlignment.CENTER);    //выравнивание по горизонтали по центру
        cellStylesMap.put("t12alC",t12alC);

        //t12alR
        XSSFCellStyle t12alR = wb.createCellStyle();
        t12alR.setFont(t12F);
        t12alR.setAlignment(HorizontalAlignment.RIGHT);    //выравнивание по горизонтали по правому краю
        cellStylesMap.put("t12alR", t12alR);

        //t12alL
        XSSFCellStyle t12alL = wb.createCellStyle();
        t12alL.setFont(t12F);
        t12alL.setAlignment(HorizontalAlignment.LEFT);    //выравнивание по горизонтали по левому краю
        cellStylesMap.put("t12alL", t12alL);

        //t12balC
        XSSFCellStyle t12balC = wb.createCellStyle();
        t12balC.setFont(t12bF);
        t12balC.setAlignment(HorizontalAlignment.CENTER);   //выравнивание по горизонтали по центру
        cellStylesMap.put("t12balC",t12balC);

        //t12balLW
        XSSFCellStyle t12balLW = wb.createCellStyle();
        t12balLW.setFont(t12bF);
        t12balLW.setAlignment(HorizontalAlignment.LEFT);    //выравнивание по горизонтали по левому краю
        t12balLW.setWrapText(true);                         //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12balLW",t12balLW);

        //t12balCW
        //Стиль для формирования списков
        XSSFCellStyle t12balCW = wb.createCellStyle();
        t12balCW.setFont(t12bF);
        t12balCW.setAlignment(HorizontalAlignment.CENTER);  //выравнивание по горизонтали по центру
        t12balCW.setWrapText(true);                         //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12balCW",t12balCW);

        //t12alCBB
        XSSFCellStyle t12alCBB = wb.createCellStyle();
        t12alCBB.setFont(t12F);
        t12alCBB.setAlignment(HorizontalAlignment.CENTER);  //выравнивание по горизонтали по центру
        t12alCBB.setBorderBottom(BorderStyle.THIN);         //Обрамление линиями снизу
        cellStylesMap.put("t12alCBB",t12alCBB);

        //t12balRB
        XSSFCellStyle t12balRB = wb.createCellStyle();
        t12balRB.cloneStyleFrom(b);
        t12balRB.setFont(t12bF);
        t12balRB.setAlignment(HorizontalAlignment.RIGHT);   //выравнивание по горизонтали по праваму краю
        cellStylesMap.put("t12balRB",t12balRB);

        //t12alCWB
        XSSFCellStyle t12alCWB = wb.createCellStyle();
        t12alCWB.cloneStyleFrom(b);
        t12alCWB.setFont(t12F);
        t12alCWB.setAlignment(HorizontalAlignment.CENTER);  //выравнивание по горизонтали по центру
        t12alCWB.setWrapText(true);                         //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12alCWB",t12alCWB);

        //t12balCWB
        XSSFCellStyle t12balCWB = wb.createCellStyle();
        t12balCWB.cloneStyleFrom(b);
        t12balCWB.setFont(t12bF);
        t12balCWB.setAlignment(HorizontalAlignment.CENTER); //выравнивание по горизонтали по центру
        t12balCWB.setWrapText(true);                        //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12balCWB",t12balCWB);

        //t12alCCWB
        //Стиль столбца Номер по порядку, Наименование курса обучения и Количество человеко-часов
        XSSFCellStyle t12alCCWB = wb.createCellStyle();
        t12alCCWB.cloneStyleFrom(b);
        t12alCCWB.setFont(t12F);
        t12alCCWB.setAlignment(HorizontalAlignment.CENTER);       //выравнивание по горизонтали по центру
        t12alCCWB.setVerticalAlignment(VerticalAlignment.CENTER); //выравнивание по вертикали по центру
        t12alCCWB.setWrapText(true);                             //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12alCCWB",t12alCCWB);

        //t12alCHWB
        XSSFCellStyle t12alCHWB = wb.createCellStyle();
        t12alCHWB.cloneStyleFrom(b);
        t12alCHWB.setFont(t12F);
        t12alCHWB.setAlignment(HorizontalAlignment.CENTER);      //выравнивание по горизонтали по центру
        t12alCHWB.setVerticalAlignment(VerticalAlignment.TOP);   //выравнивание по вертикали по верхнему краю
        t12alCHWB.setWrapText(true);                             //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12alCHWB",t12alCHWB);

        //t12alLCWB
        //Стиль столбца ФИО слушателя
        XSSFCellStyle t12alLCWB = wb.createCellStyle();
        t12alLCWB.cloneStyleFrom(b);
        t12alLCWB.setFont(t12F);
        t12alLCWB.setAlignment(HorizontalAlignment.LEFT);         //выравнивание по горизонтали по левому краю
        t12alLCWB.setVerticalAlignment(VerticalAlignment.CENTER); //выравнивание по вертикали по центру
        t12alLCWB.setWrapText(true);                              //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12alLCWB",t12alLCWB);

        //t12alLHWB
        XSSFCellStyle t12alLHWB = wb.createCellStyle();
        t12alLHWB.cloneStyleFrom(b);
        t12alLHWB.setFont(t12F);
        t12alLHWB.setAlignment(HorizontalAlignment.LEFT);         //выравнивание по горизонтали по левому краю
        t12alLHWB.setVerticalAlignment(VerticalAlignment.TOP);    //выравнивание по вертикали по верхнему краю
        t12alLHWB.setWrapText(true);                              //переноса текста в ячейке согласно ее размеру (ширине)
        cellStylesMap.put("t12alLHWB",t12alLHWB);

    }

    public XSSFCellStyle getStyle(String code){
        if (!cellStylesMap.containsKey(code)) throw new IllegalArgumentException("Имя стиля Excel не найдено [" + code +"]");
        return cellStylesMap.get(code);
    }

}
