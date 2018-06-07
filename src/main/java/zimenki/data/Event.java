package zimenki.data;

import java.util.*;

public class Event {

    private final EventTypes type;
    private final String name;
    private final String decree;
    private final DatePeriud date;
    private final List<Person> persons;
    private int sumManHours = 0;

    public Event(int type, String name, String decree, DatePeriud date) {
        this.type = EventTypes.values()[type];
        this.name = name;
        this.decree = decree;
        this.date = date;
        this.persons = new ArrayList<>();
    }

    public Event(EventTypes type, String name, String decree, DatePeriud date) {
        this.type = type;
        this.name = name;
        this.decree = decree;
        this.date = date;
        this.persons = new ArrayList<>();
    }

    public void addPerson(Person person) {
        this.persons.add(person);
        sumManHours+=person.getManHours();
    }

    public void addPersons(List<Person> persons) {
        for (Person person : persons) {
            addPerson(person);
        }
    }

    public List<Person> getPersons() {
        return persons;
    }

    public EventTypes getType() {
        return type;
    }

    public String getName() {
        return name;
    }

    public String getDecree() {
        return decree;
    }

    public DatePeriud getDate() {
        return date;
    }

    public int getNumberOfPersons(){
        return persons.size();
    }

    public int getSumManHours(){
        return sumManHours;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof Event)) return false;
        Event event = (Event) o;
        return type == event.type &&
                Objects.equals(name, event.name) &&
                Objects.equals(decree, event.decree) &&
                Objects.equals(date, event.date) &&
                Objects.equals(persons, event.persons);
    }

    @Override
    public int hashCode() {

        return Objects.hash(type, name, decree, date, persons);
    }
}