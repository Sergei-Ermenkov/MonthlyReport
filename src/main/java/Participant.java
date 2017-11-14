import java.util.Map;
import java.util.logging.Logger;

class Participant {
    private static final Logger LOGGER = Logger.getLogger(Test.class.getName());

    private String name;
    private int manHours;
    private Branches branch;


    public Participant(String name, int manHours, Branches branch) {
        this.name = name;
        this.manHours = manHours;
        this.branch = branch;
    }

    public Participant(String name, int manHours, String branch) {
        this(name, manHours, findBranch(branch));
    }

    public String getName() {
        return name;
    }

    public Branches getBranch() {
        return branch;
    }

    @Override
    public String toString() {
        return "Participant{" +
                "name='" + name + '\'' +
                ", manHours=" + manHours +
                ", branch=" + branch +
                '}';
    }

    private static Branches findBranch(String branch) {
        // делаем все буквы маленькие и удаляем пробелы.
        String branchString = branch.toLowerCase().replaceAll("\\s", "");
        // сопостовляем строчку с ключами словоря филиалов
        for (Map.Entry<String, Branches> entry : Dictionary.branchesMap.entrySet()) {
            if (branchString.contains(entry.getKey())) {
                return entry.getValue();
            }
        }
        return Branches.АДМИНИСТРАЦИЯ;
    }
}
