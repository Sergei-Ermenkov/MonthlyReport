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
........W.. - переноса текста в ячейке согласно ее размеру (ширине)
.........B. - Обрамление ячейки линиями со всех сторон
.........BB - Обрамление ячейки линиями снизу
 */
public class ExcelStyles {
    private Map<String, XSSFCellStyle> cellStylesMap = new HashMap<>();
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

        //t12alC
        XSSFCellStyle t12alC = wb.createCellStyle();
        t12alC.setFont(t12F);
        t12alC.setAlignment(HorizontalAlignment.CENTER);    //выравнивание по горизонтали по центру
        cellStylesMap.put("t12alC",t12alC);

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
        t12balRB.setFont(t12bF);
        t12balRB.setAlignment(HorizontalAlignment.RIGHT);   //выравнивание по горизонтали по праваму краю
        t12balRB.setBorderBottom(BorderStyle.THIN);         //Обрамление линиями снизу
        t12balRB.setBorderTop(BorderStyle.THIN);            //Обрамление линиями сверху
        t12balRB.setBorderLeft(BorderStyle.THIN);           //Обрамление линиями слева
        t12balRB.setBorderRight(BorderStyle.THIN);          //Обрамление линиями справа
        cellStylesMap.put("t12balRB",t12balRB);

        //t12alCWB
        XSSFCellStyle t12alCWB = wb.createCellStyle();
        t12alCWB.setFont(t12F);
        t12alCWB.setAlignment(HorizontalAlignment.CENTER);  //выравнивание по горизонтали по центру
        t12alCWB.setWrapText(true);                         //переноса текста в ячейке согласно ее размеру (ширине)
        t12alCWB.setBorderBottom(BorderStyle.THIN);         //Обрамление линиями снизу
        t12alCWB.setBorderTop(BorderStyle.THIN);            //Обрамление линиями сверху
        t12alCWB.setBorderLeft(BorderStyle.THIN);           //Обрамление линиями слева
        t12alCWB.setBorderRight(BorderStyle.THIN);          //Обрамление линиями справа
        cellStylesMap.put("t12alCWB",t12alCWB);

        //t12balCWB
        XSSFCellStyle t12balCWB = wb.createCellStyle();
        t12balCWB.setFont(t12bF);
        t12balCWB.setAlignment(HorizontalAlignment.CENTER); //выравнивание по горизонтали по центру
        t12balCWB.setWrapText(true);                        //переноса текста в ячейке согласно ее размеру (ширине)
        t12balCWB.setBorderBottom(BorderStyle.THIN);         //Обрамление линиями снизу
        t12balCWB.setBorderTop(BorderStyle.THIN);            //Обрамление линиями сверху
        t12balCWB.setBorderLeft(BorderStyle.THIN);           //Обрамление линиями слева
        t12balCWB.setBorderRight(BorderStyle.THIN);          //Обрамление линиями справа
        cellStylesMap.put("t12balCWB",t12balCWB);

        //t12alCCWB
        //Стиль столбца Номер по порядку, Наименование курса обучения и Количество человеко-часов
        XSSFCellStyle t12alCCWB = wb.createCellStyle();
        t12alCCWB.setFont(t12F);
        t12alCCWB.setAlignment(HorizontalAlignment.CENTER);       //выравнивание по горизонтали по центру
        t12alCCWB.setVerticalAlignment(VerticalAlignment.CENTER); //выравнивание по вертикали по центру
        t12alCCWB.setWrapText(true);                             //переноса текста в ячейке согласно ее размеру (ширине)
        t12alCCWB.setBorderBottom(BorderStyle.THIN);             //Обрамление линиями снизу
        t12alCCWB.setBorderTop(BorderStyle.THIN);                //Обрамление линиями сверху
        t12alCCWB.setBorderLeft(BorderStyle.THIN);               //Обрамление линиями слева
        t12alCCWB.setBorderRight(BorderStyle.THIN);              //Обрамление линиями справа
        cellStylesMap.put("t12alCCWB",t12alCCWB);

        //t12alLCWB
        //Стиль столбца ФИО слушателя
        XSSFCellStyle t12alLCWB = wb.createCellStyle();
        t12alLCWB.setFont(t12F);
        t12alLCWB.setAlignment(HorizontalAlignment.LEFT);         //выравнивание по горизонтали по левому краю
        t12alLCWB.setVerticalAlignment(VerticalAlignment.CENTER); //выравнивание по вертикали по центру
        t12alLCWB.setWrapText(true);                              //переноса текста в ячейке согласно ее размеру (ширине)
        t12alLCWB.setBorderBottom(BorderStyle.THIN);              //Обрамление линиями снизу
        t12alLCWB.setBorderTop(BorderStyle.THIN);                 //Обрамление линиями сверху
        t12alLCWB.setBorderLeft(BorderStyle.THIN);                //Обрамление линиями слева
        t12alLCWB.setBorderRight(BorderStyle.THIN);               //Обрамление линиями справа
        cellStylesMap.put("t12alLCWB",t12alLCWB);
    }

    public XSSFCellStyle getStyle(String code){
        if (!cellStylesMap.containsKey(code)) throw new IllegalArgumentException("Имя стиля Excel не найдено [" + code +"]");
        return cellStylesMap.get(code);
    }

}
