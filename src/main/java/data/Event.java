package data;

import java.util.HashSet;
import java.util.Objects;
import java.util.Set;

public class Event {

    private final EventTypes type;
    private final String name;
    private final String decree;
    private final DatePeriud date;
    private final Set<Person> persons;

    public Event(EventTypes type, String name, String decree, DatePeriud date) {
        this.type = type;
        this.name = name;
        this.decree = decree;
        this.date = date;
        this.persons = new HashSet<>();
    }

    public void addPerson(Person person) {
        this.persons.add(person);
    }

    public Set<Person> getPersons() {
        return persons;
    }

    public int personsSize(){
        return persons.size();
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