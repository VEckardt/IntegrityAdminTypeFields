package typefields.api;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;

import com.mks.api.CmdRunner;
import com.mks.api.Command;
import com.mks.api.IntegrationPoint;
import com.mks.api.IntegrationPointFactory;
import com.mks.api.Option;
import com.mks.api.Session;
import com.mks.api.response.APIConnectionException;
import com.mks.api.response.APIException;
import com.mks.api.response.ApplicationConnectionException;
import com.mks.api.response.Field;
import com.mks.api.response.Item;
import com.mks.api.response.Response;
import com.mks.api.response.WorkItem;
import com.mks.api.response.WorkItemIterator;
import java.util.TreeMap;
import typefields.utils.Html;

public class IntegritySession {

    public Map<String, WorkItem> allFields = new HashMap<>();
    public Map<String, WorkItem> allFieldsById = new HashMap<>();
    public Map<String, WorkItem> allTriggers = new HashMap<>();
    public Map<String, WorkItem> allTypes = new LinkedHashMap<>();
    public Map<String, WorkItem> allQueries = new LinkedHashMap<>();
    public Map<String, WorkItem> allCharts = new LinkedHashMap<>();
    public Map<String, WorkItem> allReports = new LinkedHashMap<>();
    public Map<String, WorkItem> allIMProjects = new LinkedHashMap<>();
    public Map<String, WorkItem> allTestObjectives = new LinkedHashMap<>();
    public Map<String, String[]> solutionProperties; // =
    // intSession.getProperties(null,
    // "ALM_MKS Solution");

    // private CmdRunner cmdRunner = null;
    private String commandUsed;
    private IntegrationPoint integrationPoint = null;
    private Session apiSession = null;

    private String user = (System.getenv("MKSSI_USER") == null ? "veckardt" : System.getenv("MKSSI_USER"));
    private String password = (System.getenv("MKSSI_PASSWORD") == null ? "password" : System.getenv("MKSSI_PASSWORD"));
    private String hostname = (System.getenv("MKSSI_HOST") == null ? "localhost" : System.getenv("MKSSI_HOST"));
    private String port = System.getenv("MKSSI_PORT") == null ? "7001" : System.getenv("MKSSI_PORT");

    public IntegritySession(Boolean blank) throws APIException {

    }

    public IntegritySession(String[] args) throws APIException {
        
        for (String arg : args) {
            if (arg.indexOf("--hostname=") == 0) {
                hostname = arg.substring("--hostname=".length(), arg.length());
            }
            if (arg.indexOf("--port=") == 0) {
                port = arg.substring("--port=".length(), arg.length());
            }
            if (arg.indexOf("--user=") == 0) {
                user = arg.substring("--user=".length(), arg.length());
            }
            if (arg.indexOf("--password=") == 0) {
                password = arg.substring("--password=".length(), arg.length());
            }
        }        

        IntegrationPointFactory ipf = IntegrationPointFactory.getInstance();
        try {

            // if (1 == 1) {
            integrationPoint = ipf.createLocalIntegrationPoint(4, 11);
            apiSession = integrationPoint.getCommonSession();
            apiSession.setDefaultHostname(hostname);
            apiSession.setDefaultUsername(user);
            apiSession.setDefaultPort(Integer.parseInt(port));
            // apiSession.setDefaultPassword(password);

            // cmdRunner = apiSession.createCmdRunner();
            // cmdRunner.setDefaultHostname(host);
            // cmdRunner.setDefaultPort(port);
            // cmdRunner.setDefaultUsername("veckardt");
            // cmdRunner.setDefaultPassword("");
            // }
            // apiSession = integrationPoint.getCommonSession();
            Command cmd = new Command(Command.IM, "connect");
            // cmd.addOption(new Option("gui"));
            this.execute(cmd);

        } catch (APIConnectionException ce) {
            System.out.println("IMConfig: Unable to connect to Integrity Server");
            throw ce;
        } catch (ApplicationConnectionException ace) {
            System.out.println("IMConfig: Integrity Client unable to connect to Integrity Server");
            throw ace;
        } catch (APIException apiEx) {
            System.out.println("IMConfig: Unable to initialize");
            throw apiEx;
        }
    }

    public Response execute(Command cmd) throws APIException {
        CmdRunner cmdRunner = apiSession.createCmdRunner();
        cmdRunner.setDefaultUsername(user);
        cmdRunner.setDefaultPassword(password);
        commandUsed = cmd.getCommandName();
        // OptionList ol = cmd.getOptionList();
        // for (int i=0; i<ol.size(); i++);
        //    Iterator o = ol.getOptions();
        //    o.

        Response response = cmdRunner.execute(cmd);
        cmdRunner.release();
        commandUsed = response.getCommandString();
        return response;
    }

    /**
     * getRelatedTypes from Relationship
     *
     * @param typeName
     * @return
     */
    public Map<String, String[]> getRelatedTypes(String typeName) {

        Map<String, String[]> relatedTypes = new TreeMap<>();

        for (WorkItem field : allFields.values()) {
            if (field.getField("type").getString().contentEquals("relationship")) {
                if (field.getField("allowedTypes") != null && field.getField("allowedTypes").getList() != null // && field.getField("description").getString().startsWith("Structural relationship")
                        ) {
                    Boolean trace = field.getField("trace").getBoolean();
                    Boolean isMultiValued = field.getField("isMultiValued").getBoolean();
                    String description = field.getField("description").getString();
                    if (!(description + "").startsWith("Structural relationship")) {
                        ListIterator it = field.getField("allowedTypes").getList().listIterator();
                        while (it.hasNext()) {
                            Item allowedTypes = (Item) (it.next());
                            if (allowedTypes.getId().contentEquals(typeName)) {
                                String toTypes = allowedTypes.getField("to").getValueAsString();
                                for (String toType : toTypes.split(",")) {
                                    relatedTypes.put(field.getId() + "+" + toType, new String[]{field.getId(), toType, trace ? "yes" : "-", isMultiValued ? "multiple allowed" : "single only", description});
                                }
                            }
                        }
                    }
                }
            }
        }

        return relatedTypes;

    }

// properties
    public Map<String, String[]> getProperties(WorkItem wi, String typeName) throws APIException {
        Map<String, String[]> properties = new LinkedHashMap<>();
        Field fld;

        if (wi == null) {
            Command cmd = new Command(Command.IM, "viewtype");
            cmd.addOption(new Option("showProperties"));
            cmd.addSelection(typeName.replace("+", " "));
            Response respo = this.execute(cmd);
            WorkItem wit = respo.getWorkItem(typeName.replace("+", " "));
            fld = wit.getField("properties");
        } else {
            fld = wi.getField("properties");
        }

        if (fld != null && fld.getList() != null) {
            @SuppressWarnings("rawtypes")
            ListIterator li = fld.getList().listIterator();
            while (li.hasNext()) {
                Item it = (Item) li.next();
                String pName = Html.toHtml(it.getField("name").getValueAsString());
                String pDescription = Html.toHtml("" + it.getField("description").getValueAsString());
                String pValue = Html.toHtml("" + it.getField("value").getValueAsString());
                properties.put(pName, new String[]{pValue, pDescription});
            }
        }
        return properties;
    }

    public Map<String, WorkItem> getIMProjects() throws APIException {
        Command cmd = new Command(Command.IM, "projects");
        // cmd.addOption(new Option("showProperties"));
        // cmd.addSelection(typeName.replace("+", " "));
        Response respo = this.execute(cmd);
        WorkItemIterator wit = respo.getWorkItems();
        while (wit.hasNext()) {
            WorkItem wi = wit.next();
            allIMProjects.put(wi.getId(), wi);
        }
        return allIMProjects;
    }

    public Map<String, WorkItem> getTestObjectives(String project) throws APIException {
        Command cmd = new Command(Command.IM, "issues");
        // im issues --query="All Test Objectives"
        // --fieldFilter=project=/Projects/Release2
        cmd.addOption(new Option("query", "All Test Objectives"));
        cmd.addOption(new Option("fieldFilter", "project=" + project));
        // cmd.addSelection(typeName.replace("+", " "));
        Response respo = this.execute(cmd);
        WorkItemIterator wit = respo.getWorkItems();
        while (wit.hasNext()) {
            WorkItem wi = wit.next();
            allTestObjectives.put(wi.getId(), wi);
        }
        return allTestObjectives;
    }

    public void setEditProperty(Map<String, String[]> properties,
            Map<String, String[]> typeEditProps, String propKey) {
        if (properties.containsKey(propKey)) {
            String field = solutionProperties.get(propKey.replace("ility.", "ilityField."))[0];
            String fieldDisplayName = this.allFields.get(field).getField("displayName")
                    .getValueAsString();
            typeEditProps.put(fieldDisplayName, new String[]{properties.get(propKey)[0], field});
        }
    }

    public List<String> getFieldList(WorkItem wi, String fieldName) {

        List<String> fieldList = new ArrayList<String>();
        Field fld = wi.getField(fieldName);
        if (fld != null && fld.getList() != null) {

            @SuppressWarnings("rawtypes")
            ListIterator li = fld.getList().listIterator();
            while (li.hasNext()) {
                Item it = (Item) li.next();
                fieldList.add(it.getId());
            }
        }
        return fieldList;
    }

    // fill the simple String map
    public Map<String, String> getFieldMap(WorkItem wi, String fieldName) {
        Map<String, String> fieldMap = new HashMap<String, String>();
        Field fld = wi.getField(fieldName);
        if (fld != null && fld.getList() != null) {
            @SuppressWarnings("rawtypes")
            ListIterator li = fld.getList().listIterator();
            while (li.hasNext()) {
                Item it = (Item) li.next();
                if (!it.getId().equals("Unspecified")) {
                    fieldMap.put(it.getId(), it.getId());
                }
            }
        }
        return fieldMap;
    }

    public void readAllObjects(String type, Map<String, WorkItem> targetMap) throws APIException {
        targetMap.clear();
        Command cmd = new Command(Command.IM, type);
        // limit the fields to make this query fast, even if someone would have 2000
        // of each
        if (type.contentEquals("triggers")) {
            cmd.addOption(new Option("fields", "name,rule,query,description,type,scriptTiming"));
        } else if (type.contentEquals("types")) {
            cmd.addOption(new Option("fields", "id,name,description"));
        } else if (type.contentEquals("queries")) {
            cmd.addOption(new Option("fields", "id,name,description,queryDefinition"));
        } else if (type.contentEquals("charts")) {
            cmd.addOption(new Option("fields", "id,name,description,query,graphStyle"));
        } else if (type.contentEquals("reports")) {
            cmd.addOption(new Option("fields", "id,name,description,query"));
        } else {
            cmd.addOption(new Option("fields", "id,name"));
        }

        Response respo = this.execute(cmd);
        // ResponseUtil.printResponse(respo, 1, System.out);
        WorkItemIterator wii = respo.getWorkItems();
        while (wii.hasNext()) {
            WorkItem wi = wii.next();
            if (type.contentEquals("reports")) {
                targetMap.put(wi.getId().replaceAll(" ", "\\+"), wi);
            } else {
                targetMap.put(wi.getId(), wi);
            }
            // if (type.contentEquals("fields"))
            // allFieldsById.put(wi.getField("id").getValueAsString(), wi);
        }
    }

    public void readAllSIObjects(String type, Map<String, WorkItem> targetMap) throws APIException {
        targetMap.clear();
        Command cmd = new Command(Command.SI, type);
        // limit the fields to make this query fast, even if someone would have 2000
        // of each

        Response respo = this.execute(cmd);
        // ResponseUtil.printResponse(respo, 1, System.out);
        WorkItemIterator wii = respo.getWorkItems();
        while (wii.hasNext()) {
            WorkItem wi = wii.next();
            targetMap.put(wi.getId(), wi);
        }
    }

    public void readAllFields(String typeCriteria) throws APIException {
        Command cmd = new Command(Command.IM, "fields");
        cmd.addOption(new Option(
                "fields",
                "ID,position,name,type,displayName,description,computation,maxLength,backedBy,default,relevanceRule,allowedTypes,trace,isMultiValued"));

        Response respo = this.execute(cmd);
        // ResponseUtil.printResponse(respo, 1, System.out);
        WorkItemIterator wii = respo.getWorkItems();
        while (wii.hasNext()) {
            WorkItem wi = wii.next();
            if (typeCriteria == null || wi.getField("type").getValueAsString().equals(typeCriteria)) {
                allFields.put(wi.getId(), wi);
                allFieldsById.put(wi.getField("id").getValueAsString(), wi);
            }
        }
    }

    // im viewfield "Shared Category"
    // im viewfield "ALM_Task Phase"
    public void readField(String fieldName, Map<String, Field> targetMap, String entryName,
            String descrField) throws APIException {
        targetMap.clear();
        Command cmd = new Command(Command.IM, "viewfield");
        cmd.addSelection(fieldName);
        // System.out.println("Reading " + fieldName + " ..");
        Response respo = this.execute(cmd);
        // ResponseUtil.printResponse(respo, 1, System.out);
        WorkItem wi = respo.getWorkItem(fieldName);
        Field picks = wi.getField(entryName);
        @SuppressWarnings("rawtypes")
        ListIterator li = picks.getList().listIterator();
        while (li.hasNext()) {
            Item it = (Item) li.next();
            targetMap.put(it.getId(), it.getField(descrField));
        }
    }

    public void release() throws APIException, IOException {
        if (1 == 2) {
            if (apiSession != null) {
                apiSession.release();
            }
            if (integrationPoint != null) {
                integrationPoint.release();
            }
        }
    }

    public String getAbout(String sectionName) {

        String result = "<hr><div style=\"font-size:x-small;white-space: nowrap;text-align:center;\">"
                + sectionName
                + "<br>Copyright &copy; 2013, 2014 PTC Inc.<br>Author: Volker Eckardt, email: veckardt@ptc.com<br>";
        Command cmd = new Command("im", "about");
        try {
            Response response = this.execute(cmd);
            WorkItem wi = response.getWorkItem("ci");
            // get the details
            result = result + wi.getField("title").getValueAsString();
            result = result + ", Version: " + wi.getField("version").getValueAsString();
            // result = result + ", Patch-Level: " +
            // wi.getField("patch-level").getValueAsString();
            result = result + ", API Version: " + wi.getField("apiversion").getValueAsString();

            return result + "</div>";
        } catch (APIException ex) {
            // Logger.getLogger(APISession.class.getName()).log(Level.SEVERE, null,
            // ex);
        } catch (NullPointerException ex) {

        }
        return result;
    }
}
