package data;

import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.Objects;

public class DatePeriud {

    private final LocalDate beginDate;
    private final LocalDate endDate;

    public DatePeriud(int month, int year) {
        this.beginDate = LocalDate.of(year, month, 1);
        this.endDate = LocalDate.of(year, month, LocalDate.of(year, month, 1).lengthOfMonth());
    }

    public DatePeriud(Row row, int beginDateCell, int endDateCell){
        beginDate = row.getCell(beginDateCell).getDateCellValue()
                .toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        endDate = row.getCell(endDateCell).getDateCellValue()
                .toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    public DatePeriud(String beginDate, String endDate){
        this.beginDate = LocalDate.parse(beginDate);
        this.endDate = LocalDate.parse(endDate);
    }

    public LocalDate getBeginDate() {
        return beginDate;
    }

    public LocalDate getEndDate() {
        return endDate;
    }

    public String getMonthYearOfBeginDate(){
        return beginDate.format(DateTimeFormatter.ofPattern("LLLL YYYY", Locale.forLanguageTag("ru"))).toLowerCase();
    }

    @Override
    public String toString() {
            return beginDate.format(DateTimeFormatter.ofPattern("dd.MM.yyyy")) + '-' + endDate.format(DateTimeFormatter.ofPattern("dd.MM.yyyy"));
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof DatePeriud)) return false;
        DatePeriud that = (DatePeriud) o;
        return Objects.equals(beginDate, that.beginDate) &&
                Objects.equals(endDate, that.endDate);
    }

    @Override
    public int hashCode() {

        return Objects.hash(beginDate, endDate);
    }
}
