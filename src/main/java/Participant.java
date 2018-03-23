import java.util.Map;
import java.util.logging.Logger;

class Participant {
    private static final Logger LOGGER = Logger.getLogger(Test.class.getName());

    private String name;
    private int manHours;
    private Branches branch;

    Participant(String name, int manHours, String branch) {
        this.name = name;
        this.manHours = manHours;
        this.branch = findBranch(branch);
    }

    String getName() {
        return name;
    }

    int getManHours() {
        return manHours;
    }

    Branches getBranch() {
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
        for (Map.Entry<String, Branches> entry : BranchDictionary.branchesMap.entrySet()) {
            if (branchString.contains(entry.getKey())) {
                return entry.getValue();
            }
        }
        return Branches.АДМИНИСТРАЦИЯ;
    }
}
