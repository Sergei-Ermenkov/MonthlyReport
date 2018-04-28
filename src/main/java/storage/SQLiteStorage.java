package storage;

import data.DatePeriud;
import data.Event;
import data.EventTypes;
import eхcel.CellData;
import data.Person;

import java.sql.*;
import java.util.*;

//todo: Не проработан вариант если мероприятие больше отчетного периуда (мероприятие включает полностью отчетный период)
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
        String sql = "SELECT DISTINCT branch  FROM report WHERE (begin_date <= ? AND end_date >= ?) AND type=?";
        try (Connection conn = this.connect();
             PreparedStatement pstmt = conn.prepareStatement(sql)) {
            pstmt.setString(1, date.getEndDate().toString());
            pstmt.setString(2, date.getBeginDate().toString());
            pstmt.setInt(3, type.ordinal());
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
        String eventsSql = "SELECT DISTINCT begin_date, end_date, event_name, decree FROM date_view WHERE type=? AND branch=? ORDER BY begin_date";
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
        String eventsSql = "SELECT DISTINCT begin_date, end_date, event_name, decree FROM date_view WHERE type=? ORDER BY begin_date";
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

    public List<Event> getEvents(DatePeriud date) throws SQLException {
        List<Event> result = new ArrayList<>();
        String dateView = createViewSqlQuery(date);
        String eventsSql = "SELECT DISTINCT begin_date, end_date, type, event_name, decree FROM date_view ORDER BY begin_date";
        String personsSql = "SELECT person_name, position, branch, man_hours FROM date_view WHERE event_name=?";
        try (Connection conn = this.connect()) {
            Statement viewPstmt = conn.createStatement();
            viewPstmt.executeUpdate(dateView);
            viewPstmt.close();

            PreparedStatement pstmt = conn.prepareStatement(eventsSql);
            ResultSet rse = pstmt.executeQuery();
            while (rse.next()) {
                result.add(new Event(rse.getInt("type"), rse.getString("event_name"), rse.getString("decree"),
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
        return "CREATE TEMP VIEW date_view AS SELECT * FROM report WHERE begin_date <='" +
                date.getEndDate().toString() +
                "' AND end_date >= '" +
                date.getBeginDate().toString() +
                "'";
    }

    private final String header = "consolidated_header";
    private final int beginBodyRow = 11;
    private final String footer = "consolidated_footer";

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
            case "consolidated_header":
                sql = "SELECT row, cell, value, style FROM consolidated_header";
                break;
            case "consolidated_footer":
                sql = "SELECT row, cell, value, style FROM consolidated_footer";
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