package storage;

import data.DatePeriud;
import data.Event;
import data.EventTypes;
import eхcel.CellData;
import data.Person;

import java.sql.*;
import java.util.*;

//todo несколько записей можно добавлять с помощю BEGIN TRANSACTION; --> в случае неудачи ROLLBACK; -->COMMIT;
//todo или batch функции
//todo при использовании транзакции драйвер использует 1 соединение не прерывая его пока транзакция не будет закрыта
//todo посмотреть на http://zametkinapolyah.ru/zametki-o-mysql/chast-9-3-komanda-rollback-v-bazax-dannyx-sqlite-operator-rollback-v-sqlite3.html#__ROLLBACK_TRANSACTION_SQLite3
public class SQLiteStorage {

    private Connection connect() throws SQLException {
        String url = "jdbc:sqlite:data.db";
        Connection conn;
        Properties properties = new Properties();
        properties.setProperty("foreign_keys", "true");
        conn = DriverManager.getConnection(url, properties);
        return conn;
    }

    public void addEventPersons(Event event) throws SQLException {
        String insert = "INSERT INTO report(begin_date, end_date, type, event_name, decree, branch, person_name, position, man_hours) VALUES(?,?,?,?,?,?,?,?,?)";

        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(insert)) {
            for (Person person : event.getPersons()) {
                pstmt.setString(1, event.getDate().getBeginDate().toString());
                pstmt.setString(2, event.getDate().getEndDate().toString());
                pstmt.setInt(3, event.getType().ordinal());
                pstmt.setString(4, event.getName());
                pstmt.setString(5, event.getDecree());
                pstmt.setString(6, person.getBranch());
                pstmt.setString(7, person.getName());
                pstmt.setString(8, person.getPosition());
                pstmt.setInt(9, person.getManHours());

                pstmt.executeUpdate();
            }
        }
    }

    /**
     * Загружает словать филиалов
     *
     * @return Словать филиалов
     * @throws SQLException Ошибки при обращении к SQLite
     */
    public Map<String, String> getBranchPatterns() throws SQLException {
        String sql = "SELECT branch_patterns.pattern, branches.name FROM branch_patterns JOIN branches ON branch_patterns.branch_id=branches.id";
        Map<String, String> branchPatterns = new HashMap<>();

        try (Connection conn = this.connect();
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                branchPatterns.put(rs.getString(1), rs.getString(2));
            }
        }
        return branchPatterns;
    }

    //Списков филиалов вошедших в отчет за указанные даты по указанному типу(отчет по обучению)
    public Set<String> getBranchesStrings(DatePeriud date, EventTypes type) throws SQLException {
        Set<String> result = new HashSet<>();
        String sql = "SELECT DISTINCT branch  FROM report WHERE ((begin_date BETWEEN ? AND ?) OR (end_date BETWEEN ? AND ?)) AND type=?";
        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, date.getBeginDate().toString());
            pstmt.setString(2, date.getEndDate().toString());
            pstmt.setString(3, date.getBeginDate().toString());
            pstmt.setString(4, date.getEndDate().toString());
            pstmt.setInt(5, type.ordinal());
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {
                result.add(rs.getString(1));
            }
            rs.close();
        }
        return result;
    }

    public List<Event> getEvents(DatePeriud date, EventTypes type, String branch) throws SQLException {
        List<Event> result = new ArrayList<>();
        String dateView = createViewSqlQuery(date);
        String eventsSql = "SELECT DISTINCT begin_date, end_date, event_name, decree FROM date_view WHERE type=? AND branch=?";
        String personsSql = "SELECT person_name, position, branch, man_hours FROM date_view WHERE branch=? AND event_name=?";
        try (Connection conn = this.connect()) {
            Statement viewPstmt = conn.createStatement();
            viewPstmt.executeUpdate(dateView);
            viewPstmt.close();

            PreparedStatement pstmt = conn.prepareStatement(eventsSql);
            pstmt.setInt(1, type.ordinal());
            pstmt.setString(2, branch);
            ResultSet rse = pstmt.executeQuery();
            while (rse.next()) {
                result.add(new Event(type, rse.getString("event_name"), rse.getString("decree"),
                        new DatePeriud(rse.getString("begin_date"), rse.getString("end_date"))));
            }
            rse.close();

            pstmt = conn.prepareStatement(personsSql);
            for (Event event : result) {
                pstmt.setString(1, branch);
                pstmt.setString(2, event.getName());
                ResultSet rsp = pstmt.executeQuery();
                while (rsp.next()) {
                    event.addPerson(new Person(rsp.getString("person_name"),
                            rsp.getString("position"),
                            rsp.getString("branch"),
                            rsp.getInt("man_hours")));
                }
                rsp.close();
            }
            pstmt.close();
        }
        return result;
    }

    public List<Event> getEvents(DatePeriud date, EventTypes type) throws SQLException {
        List<Event> result = new ArrayList<>();
        String dateView = createViewSqlQuery(date);
        String eventsSql = "SELECT DISTINCT begin_date, end_date, event_name, decree FROM date_view WHERE type=?";
        String personsSql = "SELECT person_name, position, branch, man_hours FROM date_view WHERE event_name=?";
        try (Connection conn = this.connect()) {
            Statement viewPstmt = conn.createStatement();
            viewPstmt.executeUpdate(dateView);
            viewPstmt.close();

            PreparedStatement pstmt = conn.prepareStatement(eventsSql);
            pstmt.setInt(1, type.ordinal());
            ResultSet rse = pstmt.executeQuery();
            while (rse.next()) {
                result.add(new Event(type, rse.getString("event_name"), rse.getString("decree"),
                        new DatePeriud(rse.getString("begin_date"), rse.getString("end_date"))));
            }
            rse.close();

            pstmt = conn.prepareStatement(personsSql);
            for (Event event : result) {
                pstmt.setString(1, event.getName());
                ResultSet rsp = pstmt.executeQuery();
                while (rsp.next()) {
                    event.addPerson(new Person(rsp.getString("person_name"),
                            rsp.getString("position"),
                            rsp.getString("branch"),
                            rsp.getInt("man_hours")));
                }
                rsp.close();
            }
            pstmt.close();
        }
        return result;
    }

    private String createViewSqlQuery(DatePeriud date) {
        return new StringBuilder()
                .append("CREATE TEMP VIEW date_view AS SELECT * FROM report WHERE (begin_date BETWEEN '")
                .append(date.getBeginDate().toString())
                .append("' AND '")
                .append(date.getEndDate().toString())
                .append("') OR (end_date BETWEEN '")
                .append(date.getBeginDate().toString())
                .append("' AND '")
                .append(date.getEndDate().toString())
                .append("')")
                .toString();
    }

    //(отчет по обучению, списки)
    public List<CellData> getTemplate(String name) throws SQLException {
        String sql;
        switch (name) {
            case "report_header":
                sql = "SELECT row, cell, value, style FROM report_header";
                break;
            case "list_header":
                sql = "SELECT row, cell, value, style FROM list_header";
                break;
            case "footer":
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
                        rs.getInt(2), rs.getString(3), rs.getString(4)));
            }
        }
        return result;
    }

}