package zimenki.storage;

import zimenki.data.DatePeriud;
import zimenki.data.Event;
import zimenki.data.EventTypes;
import zimenki.eхcel.CellData;
import zimenki.data.Person;

import java.sql.*;
import java.util.*;

public class SQLiteStorage {

    private Connection connect() throws SQLException {
        String url = "jdbc:sqlite:data.db";
        Connection conn;
        Properties properties = new Properties();
        properties.setProperty("foreign_keys", "true");
        conn = DriverManager.getConnection(url, properties);
        return conn;
    }

    void addEventPersons(Event event, DatePeriud report_periud) throws SQLException {
        String insert = "INSERT INTO report(report_date, begin_date, end_date, type, event_name, decree, branch, person_name, position, man_hours) VALUES(?,?,?,?,?,?,?,?,?,?)";

        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(insert)) {
            for (Person person : event.getPersons()) {
                pstmt.setString(1, report_periud.getBeginDate().toString());
                pstmt.setString(2, event.getDate().getBeginDate().toString());
                pstmt.setString(3, event.getDate().getEndDate().toString());
                pstmt.setInt(4, event.getType().ordinal());
                pstmt.setString(5, event.getName());
                pstmt.setString(6, event.getDecree());
                pstmt.setString(7, person.getBranch());
                pstmt.setString(8, person.getName());
                pstmt.setString(9, person.getPosition());
                pstmt.setInt(10, person.getManHours());

                pstmt.executeUpdate();
            }
        }
    }

    Map<String, String> getBranchPatterns() throws SQLException {
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

    public Set<String> getBranchesStrings(DatePeriud report_periud, EventTypes type) throws SQLException {
        Set<String> result = new HashSet<>();
        String sql = "SELECT DISTINCT branch FROM report WHERE report_date=? AND type=?";
        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, report_periud.getBeginDate().toString());
            pstmt.setInt(2, type.ordinal());
            ResultSet rs = pstmt.executeQuery();
            while (rs.next()) {
                result.add(rs.getString(1));
            }
            rs.close();
        }
        return result;
    }

    public List<Event> getEvents(DatePeriud report_periud, EventTypes type, String branch) throws SQLException {
        List<Event> result = new ArrayList<>();
        String eventsSql = "SELECT DISTINCT begin_date, end_date, event_name, decree FROM report WHERE report_date=? AND type=? AND branch=? ORDER BY begin_date";
        String personsSql = "SELECT person_name, position, branch, man_hours FROM report WHERE report_date=? AND branch=? AND event_name=?";
        try (Connection conn = this.connect();
             PreparedStatement pstmtEvent = conn.prepareStatement(eventsSql);
             PreparedStatement pstmtPerson = conn.prepareStatement(personsSql)) {

            pstmtEvent.setString(1, report_periud.getBeginDate().toString());
            pstmtEvent.setInt(2, type.ordinal());
            pstmtEvent.setString(3, branch);
            ResultSet rse = pstmtEvent.executeQuery();
            while (rse.next()) {
                result.add(new Event(type, rse.getString("event_name"), rse.getString("decree"),
                        new DatePeriud(rse.getString("begin_date"), rse.getString("end_date"))));
            }
            rse.close();

            for (Event event : result) {
                pstmtPerson.setString(1, report_periud.getBeginDate().toString());
                pstmtPerson.setString(2, branch);
                pstmtPerson.setString(3, event.getName());
                ResultSet rsp = pstmtPerson.executeQuery();
                while (rsp.next()) {
                    event.addPerson(new Person(rsp.getString("person_name"),
                            rsp.getString("position"),
                            rsp.getString("branch"),
                            rsp.getInt("man_hours")));
                }
                rsp.close();
            }
        }
        return result;
    }

    public List<Event> getEvents(DatePeriud report_periud, EventTypes type) throws SQLException {
        List<Event> result = new ArrayList<>();
        String eventsSql = "SELECT DISTINCT begin_date, end_date, event_name, decree FROM report WHERE report_date=? AND type=? ORDER BY begin_date";
        String personsSql = "SELECT person_name, position, branch, man_hours FROM report WHERE report_date=? AND event_name=?";
        try (Connection conn = this.connect();
             PreparedStatement pstmtEvent = conn.prepareStatement(eventsSql);
             PreparedStatement pstmtPerson = conn.prepareStatement(personsSql)) {

            pstmtEvent.setString(1, report_periud.getBeginDate().toString());
            pstmtEvent.setInt(2, type.ordinal());
            ResultSet rse = pstmtEvent.executeQuery();
            while (rse.next()) {
                result.add(new Event(type, rse.getString("event_name"), rse.getString("decree"),
                        new DatePeriud(rse.getString("begin_date"), rse.getString("end_date"))));
            }
            rse.close();

            for (Event event : result) {
                pstmtPerson.setString(1, report_periud.getBeginDate().toString());
                pstmtPerson.setString(2, event.getName());
                ResultSet rsp = pstmtPerson.executeQuery();
                while (rsp.next()) {
                    event.addPerson(new Person(rsp.getString("person_name"),
                            rsp.getString("position"),
                            rsp.getString("branch"),
                            rsp.getInt("man_hours")));
                }
                rsp.close();
            }
        }
        return result;
    }

    public List<Event> getEvents(DatePeriud report_periud) throws SQLException {
        List<Event> result = new ArrayList<>();
        String eventsSql = "SELECT DISTINCT begin_date, end_date, type, event_name, decree FROM report WHERE report_date=? ORDER BY begin_date";
        String personsSql = "SELECT person_name, position, branch, man_hours FROM report WHERE report_date=? AND event_name=?";
        try (Connection conn = this.connect();
             PreparedStatement pstmtEvent = conn.prepareStatement(eventsSql);
             PreparedStatement pstmtPerson = conn.prepareStatement(personsSql)) {

            pstmtEvent.setString(1, report_periud.getBeginDate().toString());
            ResultSet rse = pstmtEvent.executeQuery();
            while (rse.next()) {
                result.add(new Event(rse.getInt("type"), rse.getString("event_name"), rse.getString("decree"),
                        new DatePeriud(rse.getString("begin_date"), rse.getString("end_date"))));
            }
            rse.close();

            for (Event event : result) {
                pstmtPerson.setString(1, report_periud.getBeginDate().toString());
                pstmtPerson.setString(2, event.getName());
                ResultSet rsp = pstmtPerson.executeQuery();
                while (rsp.next()) {
                    event.addPerson(new Person(rsp.getString("person_name"),
                            rsp.getString("position"),
                            rsp.getString("branch"),
                            rsp.getInt("man_hours")));
                }
                rsp.close();
            }
        }
        return result;
    }

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
            case "consolidated_header":
                sql = "SELECT row, cell, value, style FROM consolidated_header";
                break;
            case "consolidated_footer":
                sql = "SELECT row, cell, value, style FROM consolidated_footer";
                break;
            default:
                throw new IllegalArgumentException("Неверный запрос шаблона из базы");
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