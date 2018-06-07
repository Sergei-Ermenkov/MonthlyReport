package zimenki.data;

import java.util.Objects;

public class Person {
    private final String name;
    private final String position;
    private final String branch;
    private final int manHours;

    public Person(String name, String position, String branch, int manHours) {
        this.name = name;
        this.position = position;
        this.branch = branch;
        this.manHours = manHours;
    }

    public String getName() {
        return name;
    }

    public String getPosition() {
        return position;
    }

    public String getBranch() {
        return branch;
    }

    public int getManHours() {
        return manHours;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof Person)) return false;
        Person person = (Person) o;
        return manHours == person.manHours &&
                Objects.equals(name, person.name) &&
                Objects.equals(position, person.position) &&
                Objects.equals(branch, person.branch);
    }

    @Override
    public int hashCode() {

        return Objects.hash(name, position, branch, manHours);
    }
}
