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


    private static void createCell(XSSFRow row, int column, String value, XSSFCellStyle cellStyle) {
        XSSFCell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }

    private static void createCell(XSSFRow row, int column, int value, XSSFCellStyle cellStyle) {
        XSSFCell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(cellStyle);
    }

    private static void createCellFormula(XSSFRow row, int column, String value, XSSFCellStyle cellStyle) {
        XSSFCell cell = row.createCell(column);
        cell.setCellFormula(value);
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
                    //todo добавить коррекцию входных значений
                    //todo добавить возможность добавления пустой строки вместо даты
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
    }


    void writeToExcel(String file, String date) throws IOException {
        if (report.isEmpty()) throw new NullPointerException("Не загружены данные в report");

        XSSFWorkbook workbook = new XSSFWorkbook();
        CellStylesList cellStylesList = new CellStylesList(workbook);

        for (Map.Entry<Branches, Map<String, List<Participant>>> filialList : report.entrySet()) {

            XSSFSheet sheet = workbook.createSheet(filialList.getKey().getFullName());
            List<List<CellData>> layoutSheet = new ArrayList<>();

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
            XSSFRow row = sheet.createRow(0);
            createCell(row, 3, "Приложение № 2.2 ", cellStylesList.t9alRStyle);
            row = sheet.createRow(1);
            createCell(row, 3, "Методики бухгалтерского и налогового учета операций, связанных с", cellStylesList.t9alRStyle);
            row = sheet.createRow(2);
            createCell(row, 3, "деятельностью УПЦ", cellStylesList.t9alRStyle);
            row = sheet.createRow(3);
            if (filialList.getKey() == Branches.АДМИНИСТРАЦИЯ) {
                createCell(row, 0, "В Администрация OOO «Газпром трансгаз Москва»", cellStylesList.t12balLWStyle);
            } else {
                createCell(row, 0, "В филиал " + filialList.getKey().getFullName(), cellStylesList.t12balLWStyle);
            }
            row = sheet.createRow(5);
            createCell(row, 0, "Отчет по обучению в УЧ (Зименки) УПЦ", cellStylesList.t12balCStyle);
            row = sheet.createRow(6);
            createCell(row, 0, "за " + date + " года", cellStylesList.t12balCStyle);
            row = sheet.createRow(8);
            createCell(row, 0, "№ п/п", cellStylesList.t12balCWBStyle);
            createCell(row, 1, "Ф.И.О. слушателя", cellStylesList.t12balCWBStyle);
            createCell(row, 2, "Наименование курса обучения", cellStylesList.t12balCWBStyle);
            createCell(row, 3, "Количество человеко-часов", cellStylesList.t12balCWBStyle);
            row = sheet.createRow(9);
            createCell(row, 0, "1", cellStylesList.t12alCWBStyle);
            createCell(row, 1, "2", cellStylesList.t12alCWBStyle);
            createCell(row, 2, "3", cellStylesList.t12alCWBStyle);
            createCell(row, 3, "4", cellStylesList.t12alCWBStyle);


            // Добавление людей по филиалам и мероприятиям
            row = sheet.createRow(10);
            int num = 0;
            for (Map.Entry<String, List<Participant>> event : filialList.getValue().entrySet()) {
                int firstRowMerged = 10 + num;
                createCell(row, 2, event.getKey(), cellStylesList.topicStyle);

                //TODO тест
                float a = row.getHeight();
                LOGGER.log(Level.FINEST, "Высота ряда с ячейкой название мероприятия: {0}", a);

                for (Participant participant : event.getValue()) {
                    createCell(row, 0, String.valueOf(++num), cellStylesList.numStyle);
                    createCell(row, 1, participant.getName(), cellStylesList.fioStyle);
                    createCell(row, 3, participant.getManHours(), cellStylesList.timeStyle);

                    row = sheet.createRow(sheet.getLastRowNum() + 1);
                }
                if (event.getValue().size() > 1) {
                    //todo выровнить ячейку по тексту
                    sheet.addMergedRegion(new CellRangeAddress(firstRowMerged, num + 9, 2, 2));
                }

            }
            // Обединение ячеек итога
            sheet.addMergedRegion(new CellRangeAddress(num + 10, num + 10, 0, 2));

            //Добавление Итога и хвоста страницы
            createCell(row, 0, "ИТОГО:", cellStylesList.t12balRBStyle);
            createCell(row, 1, "", cellStylesList.t12balRBStyle);
            createCell(row, 2, "", cellStylesList.t12balRBStyle);
            createCellFormula(row, 3, "SUM(D11:D"+ String.valueOf(num + 10) +")", cellStylesList.t12balRBStyle);
            row = sheet.createRow(sheet.getLastRowNum() + 2);
            createCell(row, 0, "Начальник УПЦ", cellStylesList.t12Style);
            createCell(row, 2, "____________________", cellStylesList.t12alCStyle);
            createCell(row, 3, "Н.В.Судак", cellStylesList.t12alCBBStyle);
            row = sheet.createRow(sheet.getLastRowNum() + 1);
            createCell(row, 2, "подпись", cellStylesList.t9alCStyle);
            createCell(row, 3, "расшифровка подписи", cellStylesList.t9alCStyle);
            //Удалена подпись Пушкова
            /*
            row = sheet.createRow(sheet.getLastRowNum() + 2);
            createCell(row, 0, "Заместитель начальника службы", cellStylesList.t12Style);
            row = sheet.createRow(sheet.getLastRowNum() + 1);
            createCell(row, 0, "по эксплуатации зданий и сооружений УПЦ", cellStylesList.t12Style);
            createCell(row, 2, "____________________", cellStylesList.t12alCStyle);
            createCell(row, 3, "А.В.Пушков", cellStylesList.t12alCBBStyle);
            row = sheet.createRow(sheet.getLastRowNum() + 1);
            createCell(row, 2, "подпись", cellStylesList.t9alCStyle);
            createCell(row, 3, "расшифровка подписи", cellStylesList.t9alCStyle);
            */
            row = sheet.createRow(sheet.getLastRowNum() + 2);
            createCell(row, 0, "Исполнитель:", cellStylesList.t12Style);
            row = sheet.createRow(sheet.getLastRowNum() + 1);
            createCell(row, 0, "Ведущий инженер АСУ ТП", cellStylesList.t12unStyle);
            createCell(row, 2, "____________________", cellStylesList.t12alCStyle);
            createCell(row, 3, "С.А. Ерменков", cellStylesList.t12alCBBStyle);
            row = sheet.createRow(sheet.getLastRowNum() + 1);
            createCell(row, 0, "должность", cellStylesList.t9alLStyle);
            createCell(row, 2, "подпись", cellStylesList.t9alCStyle);
            createCell(row, 3, "расшифровка подписи", cellStylesList.t9alCStyle);

            //Добавляем свойство "Разместить не более чем на 1 странице"
            XSSFPrintSetup printSetup = sheet.getPrintSetup();
            sheet.setFitToPage(true);
            printSetup.setFitHeight((short)1);
            printSetup.setFitWidth((short)1);

//            for (int i = 0; i < layoutSheet.size(); i++) {
//                if (layoutSheet.get(i).isEmpty()) continue;
//                XSSFRow row = sheet.createRow(i);
//                for (CellData rd : layoutSheet.get(i)) {
//                    createCell(row, rd.getNumColumn(), rd.getValue(), rd.getStyle());
//                }
//            }
            LOGGER.log(Level.FINEST, "Создан лист: {0}", filialList.getKey().getFullName());
        }

        //Запись в файл
        //TODO отключена запись в файл
        try (FileOutputStream fileOut = new FileOutputStream(file)) {
            workbook.write(fileOut);
        } catch (Exception e) {
            throw e;
        }
    }
}
