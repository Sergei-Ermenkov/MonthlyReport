import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ExcelData {
    private static final Logger LOGGER = Logger.getLogger(Test.class.getName());

    private List<Event> events = new ArrayList<>();
    private Map<Branches, Map<String, List<Participant>>> report = new HashMap<>();
    private XSSFWorkbook workbook = new XSSFWorkbook();


    private static void createCell(XSSFRow row, int column, String value, XSSFCellStyle cellStyle) {
        XSSFCell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }

    void readFromExcel(String file) throws IOException {
        String data;
        String topic;
        String name;
        TypeOfEvent type;
        String branch;
        int manHours;

        try (FileInputStream fileInputStream = new FileInputStream(file);
             XSSFWorkbook myExcelBook = new XSSFWorkbook(fileInputStream)) {
            // TODO Загрузка данных из xlsx файла.
            // Загрузка книги Excel

            // Проход по всем листам, строкам и полям.
            for (Sheet sheet : myExcelBook) {
                Event event = null;
                List<Participant> participants = new ArrayList<>();
                for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
                    Row row = sheet.getRow(i);
                    //todo продумать систему отображения ошибок о пустых значениях
                    if (row.getCell(0) == null ||
                            row.getCell(1) == null ||
                            row.getCell(2) == null ||
                            "".equals(row.getCell(0).toString()) ||
                            "".equals(row.getCell(0).toString()) ||
                            "".equals(row.getCell(0).toString())) {
                        throw new IllegalArgumentException(
                                "Пустое значение поля в листе:" + sheet.getSheetName() + ", строке:" + (i + 1));
                    }

                    if (i == 0) {
                        data = row.getCell(0).getStringCellValue();
                        topic = row.getCell(1).getStringCellValue();
                        type = Enum.valueOf(TypeOfEvent.class, row.getCell(2).getStringCellValue());
                        event = new Event(topic, data, type);
                        continue;
                    }
                    name = row.getCell(0).getStringCellValue();
                    branch = row.getCell(1).getStringCellValue();
                    manHours = (int) row.getCell(2).getNumericCellValue();
                    participants.add(new Participant(name, manHours, branch));
                    System.out.println();
                }
                event.setParticipants(participants);
                events.add(event);
            }
        } catch (Exception e) {
            throw e;
        }
    }


    void makeReport() {

        if (events.isEmpty()) throw new NullPointerException("Не загружены данные в events");

        for (Event event : events) {
            if (event.getType() == TypeOfEvent.ОБУЧЕНИЕ) {
                for (Participant participant : event.getParticipants()) {
                    List<Participant> participantList = new ArrayList<Participant>() {{
                        add(participant);
                    }};
                    Map<String, List<Participant>> map = new HashMap<String, List<Participant>>() {
                        {
                            put(event.getTopic(), participantList);
                        }
                    };

                    if (report.putIfAbsent(participant.getBranch(), map) != null) {
                        if (report.get(participant.getBranch()).putIfAbsent(event.getTopic(), participantList) != null) {
                            report.get(participant.getBranch()).get(event.getTopic()).add(participant);
                        }
                    }
                }
            }

        }
        //LOGGER.log(Level.FINEST, "NN");
    }


    void writeToExcel(String file) throws IOException {
        if (report.isEmpty()) throw new NullPointerException("Не загружены данные в report");



        for (Map.Entry<Branches, Map<String, List<Participant>>> filialList : report.entrySet()) {
            XSSFSheet sheet = workbook.createSheet(filialList.getKey().getFullName());

            // Установка ширины колонок на листе
            sheet.setColumnWidth(0, 1792);
            sheet.setColumnWidth(1, 9069);
            sheet.setColumnWidth(2, 7350);
            sheet.setColumnWidth(3, 4717);

            // Обеденение ячеек
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 2));
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 3));
            sheet.addMergedRegion(new CellRangeAddress(6, 6, 0, 3));

//            XSSFCellStyle style = workbook.createCellStyle();
            XSSFCellStyle style = new XSSFCellStyle(new StylesTable());

            XSSFRow row = sheet.createRow(0);
            createCell(row, 3, "Приложение № 2.2 ", ExcelStyle.getT9alRStyle(workbook));



//            shapka = ((1, 4, 'Приложение № 2.2 ', styT9alR),
//            (2, 4, 'Методики бухгалтерского и налогового учета операций, связанных с', styT9alR),
//            (3, 4, 'деятельностью УПЦ', styT9alR),
//            (4, 1, 'В филиал _______', styT12balLW),
//            (6, 1, 'Отчет по обучению в УЧ (Зименки) УПЦ', styT12balC),
//            (7, 1, 'за _______ 2015 года', styT12balC),
//            (9, 1, '№ п/п', styT12balCWB),
//            (9, 2, 'Ф.И.О. слушателя', styT12balCWB),
//            (9, 3, 'Наименование курса обучения', styT12balCWB),
//            (9, 4, 'Количество человеко-часов', styT12balCWB),
//            (10, 1, '1', styT12alCWB),
//            (10, 2, '2', styT12alCWB),
//            (10, 3, '3', styT12alCWB),
//            (10, 4, '4', styT12alCWB))


            LOGGER.log(Level.FINEST, "Создан лист: {0}", filialList.getKey().getFullName());
        }

        /*
         public static void main(String[] args)  throws Exception {
        Workbook wb = new XSSFWorkbook(); //or new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow((short) 2);
        row.setHeightInPoints(30);
        createCell(wb, row, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_BOTTOM);
        createCell(wb, row, (short) 1, CellStyle.ALIGN_CENTER_SELECTION, CellStyle.VERTICAL_BOTTOM);
        createCell(wb, row, (short) 2, CellStyle.ALIGN_FILL, CellStyle.VERTICAL_CENTER);
        createCell(wb, row, (short) 3, CellStyle.ALIGN_GENERAL, CellStyle.VERTICAL_CENTER);
        createCell(wb, row, (short) 4, CellStyle.ALIGN_JUSTIFY, CellStyle.VERTICAL_JUSTIFY);
        createCell(wb, row, (short) 5, CellStyle.ALIGN_LEFT, CellStyle.VERTICAL_TOP);
        createCell(wb, row, (short) 6, CellStyle.ALIGN_RIGHT, CellStyle.VERTICAL_TOP);
        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("xssf-align.xlsx");
        wb.write(fileOut);
        fileOut.close();
}


        private static void createCell(Workbook wb, Row row, short column, short halign, short valign) {
            Cell cell = row.createCell(column);
            cell.setCellValue("Align It");
            CellStyle cellStyle = wb.createCellStyle();
            cellStyle.setAlignment(halign);
            cellStyle.setVerticalAlignment(valign);
            cell.setCellStyle(cellStyle);
        }
         */

        //TODO Убрать коментарии с записи файла
        try (FileOutputStream fileOut = new FileOutputStream(file)) {
            workbook.write(fileOut);
        } catch (Exception e) {
            throw e;
        }
    }

}
