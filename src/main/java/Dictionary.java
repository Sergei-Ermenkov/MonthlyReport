import java.util.HashMap;
import java.util.Map;

public class Dictionary {
    static Map<String, Branches> branchesMap = new HashMap<String, Branches>(){{
        put("белгородско", Branches.БЕЛГОРОД);
        put("белоусовско", Branches.БЕЛОУСОВО);
        put("брянско", Branches.БРЯНСК);
        put("воронежско", Branches.ВОРОНЕЖ);
        put("гавриловско", Branches.ГАВРИЛОВСК);
        put("голубаягорка", Branches.ГОЛУБАЯ_ГОРКА);
        put("донско", Branches.ДОН);
        put("елецко", Branches.ЕЛЕЦ);
        put("истьинско", Branches.ИСТЬЕ);
        put("инженернотехническийцентр", Branches.ИТЦ);
        put("итц", Branches.ИТЦ);
        put("инженерно-техническийцентр", Branches.ИТЦ);
        put("крюковско", Branches.КРЮКОВО);
        put("курско", Branches.КУРСК);
        put("моршанско", Branches.МОРШАНСК);
        put("московско", Branches.МОСКВА);
        put("орловско", Branches.ОРЕЛ);
        put("острогожско", Branches.ОСТРОГОЖСК);
        put("путятинско", Branches.ПУТЯТИНО);
        put("серпуховско", Branches.СЕРПУХОВ);
        put("тульско", Branches.ТУЛА);
        put("уавр", Branches.УАВР);
        put("управлениеаварийно-восстановительныхработ", Branches.УАВР);
        put("управлениеаварийновосстановительныхработ", Branches.УАВР);
        put("управлениематериально-техническогоснабженияикомплектации", Branches.УМТСИК);
        put("умтсик", Branches.УМТСИК);
        put("управлениетехнологическоготранспортаиспециальнойтехники", Branches.УТТИСТ);
        put("уттист", Branches.УТТИСТ);
        put("управлениепоэксплуатациизданийисооружений", Branches.УЭЗС);
        put("управлениеэксплуатациизданийисооружений", Branches.УЭЗС);
        put("уэзс", Branches.УЭЗС);
        put("цдир", Branches.ЦДИР);
        put("центрдиагностикииреабилитации", Branches.ЦДИР);
        put("центравтогаз", Branches.ЦЕНТРАВТОГАЗ);
    }};
}
