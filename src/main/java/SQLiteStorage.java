import eхcel.CellData;

import java.sql.*;
import java.time.LocalDate;
import java.util.*;

//todo несколько записей можно добавлять с помощю BEGIN TRANSACTION; --> в случае неудачи ROLLBACK; -->COMMIT;
//todo или batch функции
//todo при использовании транзакции драйвер использует 1 соединение не прерывая его пока транзакция не будет закрыта
//todo посмотреть на http://zametkinapolyah.ru/zametki-o-mysql/chast-9-3-komanda-rollback-v-bazax-dannyx-sqlite-operator-rollback-v-sqlite3.html#__ROLLBACK_TRANSACTION_SQLite3
class SQLiteStorage{

    private final String url;

    SQLiteStorage(String database){
        url = "jdbc:sqlite:" + database;
    }

    private Connection connect() throws SQLException {
        //String url = "jdbc:sqlite:data.db";
        Connection conn;
        try {
            Properties properties = new Properties();
            properties.setProperty("foreign_keys", "true");
            conn = DriverManager.getConnection(url, properties);
        } catch (SQLException e) {
            throw e;
        }
        return conn;
    }

    /**
     * Добавляет мероприятие в базу
     * <p>Проверяет есть ли мероприятие в базе (по совпадению полей name, decree, begin_date, end_date) если такого
     * мероприятия нет то добавляет мероприятие в базу и возвращает его id.
     *
     * @param type      Тип мероприятия
     * @param name      Название мероприятия
     * @param decree    Приказ на основании которого проводится мероприятие
     * @param beginDate Дата начала мероприятия
     * @param endDate   Дата кончания мероприятия
     * @return id добавляемого мероприятия
     * @throws SQLException Ошибки при обращении к SQLite
     */
    int addEvent(EventTypes type, String name, String decree, LocalDate beginDate, LocalDate endDate) throws SQLException {
        String check = "SELECT id FROM events WHERE name=? AND decree=? AND begin_date=? AND end_date=?";
        String insert = "INSERT INTO events (type, name, decree, begin_date, end_date) VALUES (?,?,?,?,?)";
        int eventId = 0;

        try (Connection conn = this.connect()) {
            try (PreparedStatement pstmt = conn.prepareStatement(check)) {
                pstmt.setString(1, name);
                pstmt.setString(2, decree);
                pstmt.setString(3, beginDate.toString());
                pstmt.setString(4, endDate.toString());
                ResultSet rs = pstmt.executeQuery();
                if (rs != null && rs.next()) {
                    eventId = rs.getInt(1);
                }
                if (rs != null)rs.close();
            }
            if (eventId == 0) {
                try (PreparedStatement pstmt = conn.prepareStatement(insert, Statement.RETURN_GENERATED_KEYS)) {
                    pstmt.setInt(1, type.ordinal());
                    pstmt.setString(2, name);
                    pstmt.setString(3, decree);
                    pstmt.setString(4, beginDate.toString());
                    pstmt.setString(5, endDate.toString());
                    pstmt.executeUpdate();
                    ResultSet rs = pstmt.getGeneratedKeys();
                    if (rs != null && rs.next()) {
                        eventId = rs.getInt(1);
                    }
                    if (rs != null)rs.close();
                    if (eventId == 0) {
                        throw new SQLException(
                                "При добавлении мероприятия \"" + name + "\" в таблицу events, не вернулся ID строки");
                    }
                }
            }
        }
        return eventId;
    }

    /**
     * Добавляет человека в базу
     * <p>Проверяет есть ли человек в базе (по совпадению полей name, position, branch_id) если такого человека нет то
     * добавляет человека в базу и возвращает его id.
     *
     * @param name      ФИО
     * @param position  Должность
     * @param branch_id Филиал (по словарю)
     * @return id добавляемого человека
     * @throws SQLException Ошибки при обращении к SQLite
     */
    int addPerson(String name, String position, int branch_id) throws SQLException {
        String check = "SELECT id FROM persons WHERE name=? AND position=? AND branch_id=?";
        String insert = "INSERT INTO persons(name, position, branch_id) VALUES (?,?,?)";
        int personId = 0;

        try (Connection conn = this.connect()) {
            try (PreparedStatement pstmt = conn.prepareStatement(check)) {
                pstmt.setString(1, name);
                pstmt.setString(2, position);
                pstmt.setInt(3, branch_id);
                ResultSet rs = pstmt.executeQuery();
                if (rs != null && rs.next()) {
                    personId = rs.getInt(1);
                }
                if (rs != null)rs.close();
            }
            if (personId == 0) {
                try (PreparedStatement pstmt = conn.prepareStatement(insert, Statement.RETURN_GENERATED_KEYS)) {
                    pstmt.setString(1, name);
                    pstmt.setString(2, position);
                    pstmt.setInt(3, branch_id);
                    pstmt.executeUpdate();
                    ResultSet rs = pstmt.getGeneratedKeys();
                    if (rs != null && rs.next()) {
                        personId = rs.getInt(1);
                    }
                    if (rs != null)rs.close();
                    if (personId == 0) {
                        throw new SQLException(
                                "При добавлении человека \"" + name + "\" в таблицу persons, не вернулся ID строки");
                    }
                }
            }
            return personId;
        }
    }

    /**
     * Добавляет присутствие человека на мероприятии
     * <p>Проверяет присутствует ли человек на мероприятии (по совпадению полей event_id, person_id) если не
     * присутствует то довавляет.
     *
     * @param event_id  id мероприятия
     * @param person_id id человека
     * @param man_hours количество человекочасов
     * @throws SQLException Ошибки при обращении к SQLite
     */
    void addPersonToEvent(int event_id, int person_id, int man_hours) throws SQLException {
        String check = "SELECT COUNT(*) FROM events_persons WHERE event_id=? AND person_id=?";
        String insert = "INSERT INTO events_persons(event_id, person_id, man_hours) VALUES(?,?,?)";
        int count;

        try (Connection conn = this.connect()) {
            try (PreparedStatement pstmt = conn.prepareStatement(check)) {
                pstmt.setInt(1, event_id);
                pstmt.setInt(2, person_id);
                ResultSet rs = pstmt.executeQuery();
                rs.next();
                count = rs.getInt(1);
                rs.close();
            }
            if (count == 0) {
                try (PreparedStatement pstmt = conn.prepareStatement(insert)) {
                    pstmt.setInt(1, event_id);
                    pstmt.setInt(2, person_id);
                    pstmt.setInt(3, man_hours);
                    pstmt.executeUpdate();
                }
            }
        }

    }

    /**
     * Загружает словать филиалов
     *
     * @return Словать филиалов
     * @throws SQLException Ошибки при обращении к SQLite
     */
    Map<String, Integer> getBranchPatterns() throws SQLException {
        String sql = "SELECT pattern, branch_id FROM branch_patterns";
        Map<String, Integer> branchPatterns = new HashMap<>();

        try (Connection conn = this.connect();
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                branchPatterns.put(rs.getString(1), rs.getInt(2));
            }
        }
        return branchPatterns;
    }

    //Списков филиалов вошедших в отчет за указанные даты
    Set<String> getBranchesListInReport(LocalDate beginDate, LocalDate endDate, EventTypes type) throws SQLException {
        Set<String> result = new HashSet<>();
        String sql = "SELECT branch FROM report WHERE ((begin_date BETWEEN ? AND ?) OR (end_date BETWEEN ? AND ?)) AND type=? GROUP BY branch";
        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, beginDate.toString());
            pstmt.setString(2, endDate.toString());
            pstmt.setString(3, beginDate.toString());
            pstmt.setString(4, endDate.toString());
            pstmt.setInt(5, type.ordinal());
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {
                result.add(rs.getString(1));
            }
            rs.close();
        }
        return result;
    }

    // Список тем которые попали в отчет для заданного филиала
    List<String> getEventsByBranchListInReport(String branch,
                                               LocalDate beginDate,
                                               LocalDate endDate,
                                               EventTypes type) throws SQLException {
        //todo не лучше ли использовать set
        List<String> result = new ArrayList<>();
        String sql = "SELECT event_name FROM report WHERE ((begin_date BETWEEN ? AND ?) OR (end_date BETWEEN ? AND ?)) AND type=? AND branch=? GROUP BY event_name";
        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, beginDate.toString());
            pstmt.setString(2, endDate.toString());
            pstmt.setString(3, beginDate.toString());
            pstmt.setString(4, endDate.toString());
            pstmt.setInt(5, type.ordinal());
            pstmt.setString(6, branch);
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {
                result.add(rs.getString(1));
            }
            rs.close();
        }
        return result;
    }

    Map<String, Integer> getPersonsInEventListInReport(String branch,
                                                       String event_name,
                                                       LocalDate beginDate,
                                                       LocalDate endDate) throws SQLException {
        Map<String, Integer> result = new HashMap<>();
        String sql = "SELECT person_name, man_hours FROM report WHERE ((begin_date BETWEEN ? AND ?) OR (end_date BETWEEN ? AND ?)) AND branch=? AND event_name=?";
        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, beginDate.toString());
            pstmt.setString(2, endDate.toString());
            pstmt.setString(3, beginDate.toString());
            pstmt.setString(4, endDate.toString());
            pstmt.setString(5, branch);
            pstmt.setString(6, event_name);
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {
                result.put(rs.getString(1), rs.getInt(2));
            }
            rs.close();
        }
        return result;
    }

    List<CellData> getTemplate(String name) throws SQLException {
        String sql;
        switch (name){
            case "report_header":
                sql = "SELECT row, cell, value, style FROM report_header";
                break;
            case "report_footer":
                sql = "SELECT row, cell, value, style FROM report_footer";
                break;
                default:
                    throw new IllegalArgumentException();
        }
        List<CellData> result = new ArrayList<>();
        try (Connection conn = this.connect();
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                result.add(new CellData(rs.getInt(1),
                        rs.getInt(2), rs.getString(3),rs.getString(4)));
            }
        }
        return result;
    }
}

