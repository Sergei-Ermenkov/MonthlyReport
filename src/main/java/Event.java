import java.time.LocalDate;
import java.util.List;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

class Event {
    private static final Logger LOGGER = Logger.getLogger(Test.class.getName());

    private String topic;
    private DateOfEvent date;
    private TypeOfEvent type;
    private List<Participant> participants;

    Event(String topic, String date, TypeOfEvent type) {
        this.topic = topic;
        this.date = findData(date);
        this.type = type;
    }

    String getTopic() {
        return topic;
    }

    DateOfEvent getDate() {
        return date;
    }

    TypeOfEvent getType() {
        return type;
    }

    List<Participant> getParticipants() {
        return participants;
    }

    void setParticipants(List<Participant> participants) {
        this.participants = participants;
    }

    private static DateOfEvent findData(String date) {
        LocalDate startDate;
        LocalDate endDate;
        int monthStart;
        int dayStart;
        int yearStart;
        int monthEnd;
        int dayEnd;
        int yearEnd;
        //Убираем все пробелы
        String dateString = date.replaceAll("\\s", "");
        //создаем шаблон для деления по группам (число)(месяц)(год)
        String pattern = "(\\d?\\d).(\\d?\\d).((?:\\d{2})?\\d{2}).?(\\d?\\d)?.?(\\d?\\d)?.?((?:\\d{2})?\\d{2})?";
        Pattern r = Pattern.compile(pattern);
        Matcher m = r.matcher(dateString);
        m.find();
        //переводим найденный группы в числа
        dayStart = Integer.parseInt(m.group(1));
        monthStart = Integer.parseInt(m.group(2));
        yearStart = Integer.parseInt(m.group(3));
        if (yearStart < 100) {
            yearStart += 2000;
        }
        startDate = LocalDate.of(yearStart, monthStart, dayStart);

        if (m.group(4) != null || m.group(5) != null || m.group(6) != null) {
            dayEnd = Integer.parseInt(m.group(4));
            monthEnd = Integer.parseInt(m.group(5));
            yearEnd = Integer.parseInt(m.group(6));
            if (yearEnd < 100) {
                yearEnd += 2000;
            }
            endDate = LocalDate.of(yearEnd, monthEnd, dayEnd);
        } else {
            endDate = LocalDate.of(yearStart, monthStart, dayStart);
        }
        return new DateOfEvent(startDate, endDate);
    }

    @Override
    public String toString() {
        return "Event{" +
                "topic='" + topic + '\'' +
                '}';
    }
}