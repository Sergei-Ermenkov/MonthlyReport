import java.time.LocalDate;
class DateOfEvent {
    private LocalDate startDate;
    private LocalDate endDate;

    DateOfEvent(LocalDate startDate, LocalDate endDate) {
        this.startDate = startDate;
        this.endDate = endDate;
    }

    public LocalDate getStartDate() {
        return startDate;
    }

    public LocalDate getEndDate() {
        return endDate;
    }

    @Override
    public String toString() {
        return "DateOfEvent{" +
                "startDate=" + startDate +
                ", endDate=" + endDate +
                '}';
    }
}