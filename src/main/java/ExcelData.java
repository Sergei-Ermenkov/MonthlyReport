import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

class ExcelData {
    private static final Logger LOGGER = Logger.getLogger(Test.class.getName());

    private List<Event> events = new ArrayList<>();
    private Map<Branches, Map<String, List<Participant>>> report = new HashMap<>();

    private List<List<CellData>> rowForWorkbook;



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


    void writeToExcel(String file, String date) throws IOException {
        if (report.isEmpty()) throw new NullPointerException("Не загружены данные в report");

        XSSFWorkbook workbook = new XSSFWorkbook();
        ExcelStyles excelStyles = new ExcelStyles(workbook);

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

            // Добавление шапки в список для записи
            rowForWorkbook = new ArrayList<>();
            rowForWorkbook.add(0, new ArrayList<CellData>(){{ add(new CellData(3, "Приложение № 2.2 ", excelStyles.T9alRStyle));}});
            rowForWorkbook.add(1, new ArrayList<CellData>(){{ add(new CellData(3, "Методики бухгалтерского и налогового учета операций, связанных с", excelStyles.T9alRStyle));}});
            rowForWorkbook.add(2, new ArrayList<CellData>(){{ add(new CellData(3, "деятельностью УПЦ", excelStyles.T9alRStyle));}});
            rowForWorkbook.add(3, new ArrayList<CellData>(){{ add(new CellData(0, "В филиал _______", excelStyles.T12balLWStyle));}});
            rowForWorkbook.add(4, new ArrayList<>());
            rowForWorkbook.add(5, new ArrayList<CellData>(){{ add(new CellData(0, "Отчет по обучению в УЧ (Зименки) УПЦ", excelStyles.T12balCStyle));}});
            rowForWorkbook.add(6, new ArrayList<CellData>(){{ add(new CellData(0, "за _______ ХХХХ года", excelStyles.T12balCStyle));}});
            rowForWorkbook.add(7, new ArrayList<>());
            rowForWorkbook.add(8, new ArrayList<CellData>(){{ add(new CellData(0, "№ п/п", excelStyles.T12balCWBStyle));}});
            rowForWorkbook.get(8).add(new CellData(1, "Ф.И.О. слушателя", excelStyles.T12balCWBStyle));
            rowForWorkbook.get(8).add(new CellData(2, "Наименование курса обучения", excelStyles.T12balCWBStyle));
            rowForWorkbook.get(8).add(new CellData(3, "Количество человеко-часов", excelStyles.T12balCWBStyle));
            rowForWorkbook.add(8, new ArrayList<CellData>(){{ add(new CellData(0, "№ п/п", excelStyles.T12balCWBStyle));}});
            rowForWorkbook.get(8).add(new CellData(1, "Ф.И.О. слушателя", excelStyles.T12balCWBStyle));
            rowForWorkbook.get(8).add(new CellData(2, "Наименование курса обучения", excelStyles.T12balCWBStyle));
            rowForWorkbook.get(8).add(new CellData(3, "Количество человеко-часов", excelStyles.T12balCWBStyle));
            rowForWorkbook.add(9, new ArrayList<CellData>(){{ add(new CellData(0, "1", excelStyles.T12alCWBStyle));}});
            rowForWorkbook.get(9).add(new CellData(1, "2", excelStyles.T12alCWBStyle));
            rowForWorkbook.get(9).add(new CellData(2, "3", excelStyles.T12alCWBStyle));
            rowForWorkbook.get(9).add(new CellData(3, "4", excelStyles.T12alCWBStyle));

            // Меняет в шапке название филиала
            if (filialList.getKey() == Branches.АДМИНИСТРАЦИЯ){
                rowForWorkbook.get(3).get(0).setValue("В Администрация OOO «Газпром трансгаз Москва»");
            }else {
                rowForWorkbook.get(3).get(0).setValue("В филиал " + filialList.getKey().getFullName());
            }

            // Меняем месяц
            rowForWorkbook.get(6).get(0).setValue("за " + date + " года");

            for (Map.Entry<String, List<Participant>> event: filialList.getValue().entrySet()) {
                ArrayList<CellData> row = new ArrayList<>();

                //CellData cellData = new CellData();
            }

            //            # Добавляем людей и мероприятия в отчет
//            numPers = 0
//            for namMerop, peoples in meropr.items():
//            magerStart = 11+numPers
//            addCell(sheet, (11+numPers, 3, namMerop, styMerop))
//            for people in peoples:
//            addCell(sheet, (11+numPers, 1, str(numPers+1), styNum))
//            addCell(sheet, (11+numPers, 2, people[0], styFIO))
//            addCell(sheet, (11+numPers, 4, people[1], styTime))
//            numPers += 1
//            total_hour += people[1]
//            total_pers += 1
//            if len(peoples) > 1:
//            sheet.merge_cells(start_row=magerStart, start_column=3, end_row=10+numPers, end_column=3)

            for (int i = 0; i < rowForWorkbook.size(); i++){
                if (rowForWorkbook.get(i).isEmpty()) continue;
                XSSFRow row = sheet.createRow(i);
                for (CellData rd : rowForWorkbook.get(i)) {
                    createCell(row, rd.getNumColumn(),rd.getValue(),rd.getStyle());
                }
            }


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
