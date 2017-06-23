/*
 * Copyright:      Copyright 2015 (c) Parametric Technology GmbH
 * Product:        PTC Integrity Lifecycle Manager
 * Author:         V. Eckardt, Senior Consultant ALM
 * Purpose:        Custom Developed Code
 * **************  File Version Details  **************
 * Revision:       $Revision: 1.2 $
 * Last changed:   $Date: 2017/03/17 00:04:01CET $
 */
package typefields;

import adminexport.models.workflow.WorkflowModel;
import java.util.*;
import com.mks.api.*;
import com.mks.api.response.*;
import static com.ptc.services.common.tools.FileUtils.canWriteFile;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import static java.lang.System.out;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import static typefields.PickListValues.ListPickListValues;
import typefields.api.IntegritySession;
import typefields.utils.Html;

/**
 *
 * @author veckardt
 */
public class TypeFields extends TypeOutput {

    /**
     * @param args the command line arguments
     */
    private static final String outputRoot = "c:\\IntegrityAdminExport\\";
    private static final String templateFile = outputRoot + "Integrity_TypeFields_Template.xlsx";
    public static final String imageRoot = outputRoot + "images/";
    public static IntegritySession intSession;

    public static HashMap<String, WorkItem> relatedQueries = new LinkedHashMap<>();
    public static HashMap<String, WorkItem> relatedCharts = new LinkedHashMap<>();
    public static HashMap<String, WorkItem> relatedReports = new LinkedHashMap<>();

    /**
     *
     * @param args
     * @throws APIException
     * @throws IOException
     * @throws Exception
     */
    public static void main(String[] args) throws APIException, IOException, Exception {

        // updateProgress(1, 20);
        if (!canWriteFile(targetfile.replace("&1", "1"))) {
            // updateProgress(1, 1);
            log("ERROR: Please close the file " + targetfile + " first, can not write to it!", 1);
            return;
        }

        openExcelFile(templateFile);

        // if (args.length == 0) {
        //     Copyright.usage();
        //     System.exit(1);
        // } 
        out.println("*****************************************************");
        out.println("* Integrity Administration - Type Fields and States *");
        out.println("*****************************************************");
        // String title = "Type Fields and States - V1.0";

        // Connect to Integrity
        intSession = new IntegritySession(args);

        intSession.readAllObjects("types", allTypes);
        intSession.readAllObjects("queries", intSession.allQueries);
        intSession.readAllObjects("charts", intSession.allCharts);
        intSession.readAllObjects("reports", intSession.allReports);
        intSession.readAllObjects("triggers", intSession.allTriggers);
        intSession.readAllFields(null);

        intSession.solutionProperties = intSession.getProperties(null, "MKS Solution");

        int fileNum = 1;
        int sheetNum = 0;
        for (String typeName : allTypes.keySet()) // String typeName = "Defect";
        {

            if (sheetNum % 15 == 0 && sheetNum != 0) {
                sheetNum = 0;
                workbook.removeSheetAt(0);
                writeExcelFile(targetfile.replace("&1", "" + fileNum++));
                openExcelFile(templateFile);
            }

            log("Analysing type '" + typeName + "' ...", 1);
            rownum = 6;
            lastRowNum = -1;

            sheetNum++;
            sheet = workbook.cloneSheet(0);
            workbook.setSheetName(sheetNum, typeName.length() > 30 ? typeName.substring(0, 30) : typeName);

            if (sheetNum == 10) {
                // System.exit(1);
            }

            addPicure(typeName);

            // Build the selection list for types
            // out.println(Html.itemSelectField("type", allTypes, false, typeName, null));
            // if a type is provided or selected
            if (typeName != null && !typeName.isEmpty()) {

                // read all field details from Integrity
                // loop through all types in the document hierarchy  (segment > node > shared node)
                // while (!typeName.isEmpty()) {
                try {

                    stateTransMap.clear();
                    relatedQueries.clear();
                    relatedCharts.clear();
                    relatedReports.clear();

                    // execute the command
                    Command cmd = new Command(Command.IM, "viewtype");
                    cmd.addOption(new Option("showProperties"));
                    cmd.addSelection(typeName);

                    Response respo = intSession.execute(cmd);
                    // ResponseUtil.printResponse(respo, 1, System.out);

                    if (respo.getExitCode() == 0) {

                        WorkItem wi = respo.getWorkItem(typeName);

                        WorkflowModel wfm = new WorkflowModel();
                        wfm.display(getCharsOnly(wi.getField("Name").getValueAsString()), wi.getField("stateTransitions"));

                        String docClass = wi.getField("documentClass").getValueAsString();

                        for (String queryName : intSession.allQueries.keySet()) {
                            findTypeInQuery(intSession.allQueries.get(queryName), typeName, docClass, intSession.allQueries.get(queryName));
                        }

                        for (String chartName : intSession.allCharts.keySet()) {
                            if (relatedQueries.containsKey(intSession.allCharts.get(chartName).getField("query").getValueAsString())) {
                                relatedCharts.put(chartName, intSession.allCharts.get(chartName));
                            }
                        }
                        for (String reportName : intSession.allReports.keySet()) {
                            if (relatedQueries.containsKey(intSession.allReports.get(reportName).getField("query").getValueAsString())) {
                                relatedReports.put(reportName, intSession.allReports.get(reportName));
                            }
                        }

                        // out.println("<hr><table border=1>");
                        println(5, 1, "Name");
                        println(5, 2, typeName);
                        colorCell();
                        rownum++;
                        println(5, 1, "Description");
                        println(5, 2, wi.getField("description").getValueAsString());
                        colorCell();
                        rownum++;
                        println(5, 1, "Document Class");
                        println(5, 2, wi.getField("documentClass").getValueAsString());
                        colorCell();
                        rownum++;
                        println(5, 1, "Phase Field");
                        println(5, 2, wi.getField("phaseField").getValueAsString());
                        colorCell();
                        rownum++;
                        println(5, 1, "Associated Type");
                        println(5, 2, wi.getField("associatedType").getValueAsString());
                        colorCell();
                        rownum++;
                        println(5, 1, "View Presentation");
                        println(5, 2, wi.getField("viewPresentation").getValueAsString());
                        colorCell();
                        rownum++;
                        println(5, 1, "Modified By");
                        println(5, 2, wi.getField("modifiedBy").getValueAsString());
                        colorCell();
                        rownum++;
                        println(5, 1, "Last Modified");
                        println(5, 2, wi.getField("lastModified").getValueAsString());
                        colorCell();
                        rownum++;
                        //out.println("</table>");

                        // get the next type name in the hierarchy, or set to "blank" when done
//                            try {
//                                if (wi.getField("associatedType").getValueAsString().contentEquals("none")) {
//                                    typeName = "";
//                                } else {
//                                    typeName = wi.getField("associatedType").getValueAsString();
//                                }
//                            } catch (Exception e) {
                        // in case of invalid response, stop herewith
                        // typeName = "";
//                            }
                        Map<String, String> phases = new HashMap<>();

                        if (wi.getField("phaseField").getValueAsString() != null
                                && !wi.getField("phaseField").getValueAsString().contentEquals("null")) {
                            // build up the map for the phase field
                            cmd = new Command(Command.IM, "viewfield");
                            cmd.addSelection(wi.getField("phaseField").getValueAsString());
                            respo = intSession.execute(cmd);
                            WorkItem wph = respo.getWorkItem(wi.getField("phaseField").getValueAsString());
                            // ResponseUtil.printResponse(respo, 1, System.out);
                            Field fld = wph.getField("phases");
                            if (fld != null && fld.getList() != null) {
                                @SuppressWarnings("rawtypes")
                                ListIterator lip = fld.getList().listIterator();
                                while (lip.hasNext()) {
                                    Item it = (Item) lip.next(); // im.PhaseList.Entry
                                    Field fld2 = it.getField("states");
                                    @SuppressWarnings("rawtypes")
                                    ListIterator li2 = fld2.getList().listIterator();
                                    while (li2.hasNext()) {
                                        String it2 = (String) li2.next(); // Field
                                        phases.put(it2, it.getId());
                                    }
                                }
                            }
                        }

                        Map<String, String> stateTransHM = new HashMap<>();
                        Map<String, String> stateGroupsLHM = new LinkedHashMap<>();

                        // BEGIN: Workflow State Transitions and permitted Groups
//                        stateTrans[0] = "From State";
//                        stateTrans[1] = "To State";
//                        stateTrans[2] = "Permitted Groups";
//
//                        stateTransMap.put("1", stateTrans);
                        Field fld = wi.getField("stateTransitions");

                        if (fld != null && fld.getList() != null) {
                            ListIterator liSt = fld.getList().listIterator();
                            while (liSt.hasNext()) {
                                Item it = (Item) liSt.next();
                                String currFromState = "";
                                String fromState = it.getId();
                                fld = it.getField("targetStates");
                                @SuppressWarnings("rawtypes")
                                ListIterator li2 = fld.getList().listIterator();
                                while (li2.hasNext()) {
                                    Item it2 = (Item) li2.next();
                                    String toState = it2.getId();
                                    // if (!fromState.contentEquals(toState)) {
                                    String[] stateTrans = new String[4];
                                    if (currFromState.contentEquals(fromState)) {
                                        stateTrans[0] = "";
                                    } else {
                                        stateTrans[0] = fromState;
                                        currFromState = fromState;
                                    }
                                    if (fromState.contentEquals(toState)) {
                                        stateTrans[3] = "Yes";
                                    } else {
                                        stateTrans[3] = "";
                                    }

                                    stateTrans[1] = toState;
                                    fld = it2.getField("permittedGroups");
                                    @SuppressWarnings("rawtypes")
                                    ListIterator li3 = fld.getList().listIterator();
                                    while (li3.hasNext()) {
                                        Item it3 = (Item) li3.next();
                                        String permGroup = it3.getId();
                                        // out.println(fromState+" => "+toState+ ", permGroup: "+permGroup);
                                        stateTrans[2] = (stateTrans[2] == null ? "" : stateTrans[2] + ", ") + permGroup;
                                        stateTransHM.put(fromState + "+" + toState + "+" + permGroup, "x");
                                        stateGroupsLHM.put(permGroup, "x");
                                    }
                                    // out.println("1: "+fromState+":"+toState+" => "+stateTrans.toString());
                                    stateTransMap.put(fromState + ":" + toState, stateTrans);
                                    // }
                                }
                            }
                        }

//                            String permGrps = "";
//                            for (Object key : stateGroupsLHM.keySet()) {
//                                permGrps = permGrps + "<th>" + key.toString() + "</th>";
//                            }
                        // not yet here:  stateTrans = stateTrans.replace("<th>Permitted Groups</th>", permGrps);
                        // END: Workflow State Transitions and permitted Groups
                        // Properties
                        Map<String, String[]> properties = intSession.getProperties(wi, null);

                        //
                        typeEditProps.clear();

                        // editable Fields based on other fields
                        if (docClass.contentEquals("segment") || docClass.contentEquals("node")) {
                            // MKS.RQ.Editability.1
                            // Fields that are editable only when the field referenced by property
                            // MKS.RQ.EditabilityField.1 = "Valid Change Order" has value "true".

                            String[] editPropertyName = {"MKS.RQ.Editability.1", "MKS.RQ.Editability.Document.1", "MKS.RQ.Editability.Document.2", "MKS.RQ.Editability.Document.3"};
                            for (String editPropertyName1 : editPropertyName) {
                                intSession.setEditProperty(properties, typeEditProps, editPropertyName1);
                            }
                            // (Configurable) Fields that are editable only when the field referenced by property
                            // MKS.RQ.EditabilityField.Document.1 has value "true".
                            // Allow Edits
                            // Allow Trace = "MKS.RQ.Editability.Document.2";
                            // Allow Link = "MKS.RQ.Editability.Document.3";
                        }

                        // visibleFields
                        List<String> visibleFields = intSession.getFieldList(wi, "visibleFields");
                        // systemManagedFields
                        List<String> systemManagedFields = intSession.getFieldList(wi, "systemManagedFields");
                        // copyFields
                        Map<String, String> copyFields = intSession.getFieldMap(wi, "copyFields");
                        // significantEditFields
                        Map<String, String> significantEditFields = intSession.getFieldMap(wi, "significantEdit");
                        // stateTransitions
                        Map<String, String> stateTransitions = intSession.getFieldMap(wi, "stateTransitions");

                        // mandatoryFields
                        Map<String, String> mandatoryFields = new HashMap<>();
                        fld = wi.getField("mandatoryFields");
                        if (fld != null && fld.getList() != null) {
                            ListIterator li = fld.getList().listIterator();
                            while (li.hasNext()) {
                                Item it = (Item) li.next(); // State
                                Field fld2 = it.getField("fields");
                                @SuppressWarnings("rawtypes")
                                ListIterator li2 = fld2.getList().listIterator();
                                while (li2.hasNext()) {
                                    Item it2 = (Item) li2.next(); // Field
                                    mandatoryFields.put(it2.getId() + "+" + it.getId(), "x");
                                }
                            }
                        }

                        // Begin with the main table output
                        rownum++;
                        println(2, 1, "Fields and States:");
                        int k = 1;
                        rownum++;
                        println(3, k++, "Fields");
                        println(3, k++, "DisplayName");
                        println(3, k++, "Description");
                        println(3, k++, "Type");
                        println(3, k++, "Visible");
                        println(3, k++, "Copy");
                        println(3, k++, "Sign Edit");
                        colorRow(2);

                        for (String key : typeEditProps.keySet()) {
                            println(3, k++, key);
                        }

                        // out.println("<th>if "+edit1Field+"</th><th>if "+editDoc1Field+"</th><th>if "+editDoc2Field+"</th><th>if "+editDoc3Field+"</th>");
                        for (Object key : stateTransitions.keySet()) {
                            String phase = (String) phases.get(key.toString());
                            if (phase != null && !phase.equals("")) {
                                println(2, k++, key.toString() + " (" + phase + ")");
                            } else {
                                println(2, k++, key.toString());
                            }
                        }

                        // for visibleFields
                        // out.println(getFieldRow(visibleFields, "x", copyFields, significantEditFields, stateTransitions, mandatoryFields));
                        // for systemManagedFields
                        // out.println(getFieldRow(systemManagedFields, "S", copyFields, significantEditFields, stateTransitions, mandatoryFields));
                        // fill the arrays
                        typeFields.clear();
                        sortedTypeFields.clear();
                        getFieldRow(intSession, visibleFields, "x", copyFields, significantEditFields,
                                typeEditProps, stateTransitions, mandatoryFields);
                        getFieldRow(intSession, systemManagedFields, "S", copyFields, significantEditFields,
                                typeEditProps, stateTransitions, mandatoryFields);

                        // analyse the presentation xml
                        // havn't found a better way than SQL :)
                        if (true) {
                            String viewPresentation = wi.getField("viewPresentation").getValueAsString();
                            // out.println(">>> viewPresentation => "+viewPresentation);
                            String sql = "select Contents from DBFile where Name like '%/" + viewPresentation
                                    + ".xml'";
                            Command cmd2 = new Command(Command.IM, "diag");
                            cmd2.addOption(new Option("diag", "runsql"));
                            cmd2.addOption(new Option("param", sql));
                            Response respo2 = intSession.execute(cmd2);

                            /// out.println(respo2.toString());
                            // ResponseUtil.printResponse(respo2, 1, System.out);
                            if (respo2.getExitCode() == 0) {
                                Result result = respo2.getResult();
                                String message = result.getMessage();

                                if (message.length() > 100) {
                                    int p = message.indexOf("xml");
                                    // out.println(message);
                                    // out.println("~" + p + "~" + message.length() + "~");
                                    Document doc = loadXMLFromString(message.substring(p - 2));

                                    NodeList nodes1 = doc.getElementsByTagName("Tab");
                                    // out.println(nodes1.getLength() + " Tabs found to analyse ...");
                                    for (int h = 0; h < nodes1.getLength(); h++) {
                                        Node node1 = nodes1.item(h);
                                        // out.println(node1.getNodeName());
                                        NamedNodeMap nnm = node1.getAttributes();
                                        Node nName = nnm.getNamedItem("name");
                                        Node nPosition = nnm.getNamedItem("position");
                                        // out.println(nName.getNodeValue() + " /" + nPosition.getNodeValue() + "/");
                                        // out.println(nPosition.getNodeValue());
                                        rownum++;
                                        println(4, 1, "Tab: " + nName.getNodeValue() + "");
                                        colorCell();

                                        if (node1 instanceof Element) {
                                            Element e = (Element) node1;
                                            NodeList nl = e.getElementsByTagName("FieldValue");
                                            // out.println(nl.getLength()+" FieldValues found to analyse ...");
                                            for (int g = 0; g < nl.getLength(); g++) {
                                                Node node2 = nl.item(g);
                                                NamedNodeMap nnm2 = node2.getAttributes();
                                                Node nName2 = nnm2.getNamedItem("fieldID");
                                                // Node nPosition2 = nnm2.getNamedItem("textStyle");
                                                // out.println(nName2.getNodeValue());  // die feld id

                                                WorkItem fild = (WorkItem) intSession.allFieldsById.get(nName2
                                                        .getNodeValue());
                                                // out.println(".. " + fild.getField("name").getValueAsString() + " /" + nName2.getNodeValue() + "/");
                                                // out.println(nPosition2.getNodeValue());

                                                if (fild != null && fild.getField("name") != null) {
                                                    String fName = fild.getField("name").getValueAsString();
                                                    String[] fData = (String[]) typeFields.get(fName);
                                                    // if ((fData == null) || fData.equals("") || fData.equals("null")) {
                                                    //     out.println("<i>Note: field '"
                                                    //             + fName
                                                    //             + "' is part of the presentation template, but not a visible field</i><br>");
                                                    // } else {
                                                    //     String color = ((g % 2 == 1) ? "background-color:#f6f6f6;"
                                                    //             : "background-color:#FDFDFD;");
                                                    //     out.println(fData.replaceAll("#bgcolor#", color));
                                                    // }
                                                    if (fData != null) {
                                                        int j = 1;
                                                        rownum++;
                                                        for (String data : fData) {
                                                            println(rownum % 2, j++, data);
                                                        }
                                                        colorRow(rownum % 2);
                                                    }
                                                    sortedTypeFields.remove(fName);
                                                }
                                            }
                                        }
                                    }

                                }
                            }

                            rownum++;
                            rownum++;
                            println(2, 1, "Additional Fields:");
                            colorCell();

                            for (String[] fData : sortedTypeFields.values()) {

                                rownum++;
                                // String[] fData = sortedTypeFields.get(key);
                                int count = 1;
                                for (String data : fData) {
                                    println(rownum % 2, count++, data); // .replaceAll("#bgcolor#", color));
                                }
                            }
                        }

                        rownum++;
                        // footer with field counters
                        println(2, 4, "Count:");
                        println(2, 5, "" + (visibleFields.size() + systemManagedFields.size()));
                        println(2, 6, "" + (copyFields.size()));
                        println(2, 7, "" + (significantEditFields.size()));

                        // Allow Traces, Allow Links etc.
                        int l = 8;
                        for (String[] values : typeEditProps.values()) {
                            String propFields = values[0];
                            println(4, l++, "" + (propFields.split(",").length));
                            // out.println(typeEditProps.get(key)[1]);
                        }

                        // Legend
                        rownum++;
                        rownum++;
                        println(3, 1, "Legend:");
                        colorRow(2);

                        rownum++;
                        println(0, 1, "Visible");
                        println(0, 2, "Visible Fields, an S indicates a System Managed Field");
                        rownum++;
                        println(0, 1, "Copy");
                        println(0, 2, "Copy Fields");
                        rownum++;
                        println(0, 1, "Sign&nbsp;Edit");
                        println(0, 2, "Significant Edit Fields");
                        rownum++;
                        println(0, 1, "States");
                        println(0, 2, "Checked fields below the states are mandatory. The state order does not necessarily show the order in the workflow.");
                        rownum++;
                        println(0, 1, "Additional Fields");
                        println(0, 2, "These fields will usually show up at the bottom of the first tab, although system fields may not");

                        // Add legend for the additional conditional Edit fields
                        for (Object key : typeEditProps.keySet()) {
                            rownum++;
                            println(0, 1, "if " + key.toString());
                            println(0, 2, "Edit is only possible if the logical field '"
                                    + key + "' returns true");
                        }

                        // out.println(legend.getLegend());
                        rownum++;
                        rownum++;
                        println(3, 1, "Specific Document Fields:");
                        colorCell();
                        rownum++;
                        println(4, 1, "Field");
                        println(4, 2, "Display Name");
                        println(4, 3, "Computation");
                        colorRow(2);
                        for (String[] values : typeEditProps.values()) {
                            rownum++;
                            String propFieldName = values[1];
                            WorkItem field = intSession.allFields.get(propFieldName);
                            String displayName = field.getField("displayName").getValueAsString();
                            String Computation = field.getField("Computation").getValueAsString();
                            println(rownum % 2, 1, (propFieldName));
                            println(rownum % 2, 2, displayName);
                            println(rownum % 2, 3, Computation); // .replaceAll("\nOR\n", "<br>OR ").replaceAll("\nor\n", "<br>OR "));
                        }

                        rownum++;
                        rownum++;
                        println(3, 1, "Related Types:");
                        Map<String, String[]> relatedTypes = intSession.getRelatedTypes(typeName);
                        rownum++;
                        println(4, 1, "Relationship");
                        println(4, 2, "Target Type");
                        println(4, 3, "Description");
                        println(4, 4, "Multivalue");
                        println(4, 5, "Trace");
                        // println(4, 3, "Rule");
                        for (String[] relatedType : relatedTypes.values()) {
                            // log(relatedType[0] + " => " + relatedType[1], 2);
                            rownum++;
                            println(rownum % 2, 1, relatedType[0]);
                            println(rownum % 2, 2, relatedType[1]);
                            println(rownum % 2, 3, relatedType[4]);
                            println(rownum % 2, 5, relatedType[2]);
                            println(rownum % 2, 4, relatedType[3]);

                        }

                        rownum++;
                        rownum++;
                        println(3, 1, "Referenced Triggers:");
                        colorCell();
                        rownum++;
                        println(4, 1, "Trigger");
                        println(4, 2, "Rule");
                        println(4, 3, "Description");
                        println(4, 4, "Type");
                        println(4, 5, "Timing");
                        colorRow(2);

                        int tCnt = 0;
                        for (WorkItem wit : intSession.allTriggers.values()) {

                            // WorkItem wit = intSession.allTriggers.get(key);
                            String tName = wit.getField("name").getValueAsString();
                            String tRule = wit.getField("rule").getValueAsString();
                            String tQuery = wit.getField("query").getValueAsString();
                            String tType = wit.getField("type").getValueAsString();
                            String tTiming = wit.getField("scriptTiming").getValueAsString();
                            String tDescription = wit.getField("description").getValueAsString();
                            if (tRule != null
                                    && (tRule.contains("[\"Type\"] = \"" + wi.getField("name").getValueAsString()
                                            + "\"") || tRule.contains("is " + docClass))) {
                                rownum++;
                                tCnt++;
                                println(rownum % 2, 1, tName);
                                println(rownum % 2, 2, tRule);
                                println(rownum % 2, 3, tDescription);
                                println(rownum % 2, 4, tType);
                                println(rownum % 2, 5, tTiming);
                            }

                        }
                        if (tCnt == 0) {
                            rownum++;
                            println(0, 1, "No directly related triggers found.");
                        }

                        // Properties
                        rownum++;
                        rownum++;
                        println(3, 1, "Type Properties:");
                        colorCell();
                        if (properties.size() > 0) {
                            rownum++;
                            println(4, 1, "Property");
                            println(4, 2, "Value");
                            println(4, 3, "Description");
                            colorRow(2);
                            for (Object key : properties.keySet()) {
                                // out.println(key.toString());
                                rownum++;
                                String[] propValue = properties.get(key);
                                String pValue = propValue[0].replaceAll(",", ", ").replaceAll(";", "; ");
                                String pDescr = propValue[1];
                                println(rownum % 2, 1, key.toString());
                                println(rownum % 2, 2, pValue);
                                println(rownum % 2, 3, pDescr);
                            }

                        } else {
                            rownum++;
                            println(0, 1, "No type properties defined.");
                        }
                        rownum++;
                        rownum++;
                        // related triggers
                        // State Transitions
                        println(3, 1, "Workflow State Transitions and permitted Groups:");
                        colorCell();
                        rownum++;
                        println(4, 1, "From State");
                        println(4, 2, "To State");
                        println(4, 3, "Permitted Groups");
                        println(4, 4, "Self Transition");
                        colorRow(2);
                        for (String[] trans : stateTransMap.values()) {
                            rownum++;
                            println(rownum % 2, 1, trans[0]);
                            println(rownum % 2, 2, trans[1]);
                            println(rownum % 2, 3, trans[2]);
                            println(rownum % 2, 4, trans[3]);
                        }

                        // related queries
                        rownum++;
                        rownum++;
                        // related triggers
                        // State Transitions
                        println(3, 1, "Related Admin Queries:");
                        colorCell();
                        if (relatedQueries.size() > 0) {
                            rownum++;
                            println(4, 1, "Name");
                            println(4, 2, "-");
                            println(4, 3, "Description");
                            println(4, 4, "Is Admin");
                            colorRow(2);
                            for (WorkItem query : relatedQueries.values()) {
                                rownum++;
                                println(rownum % 2, 1, query.getId());
                                println(rownum % 2, 2, "");
                                println(rownum % 2, 3, query.getField("Description").getValueAsString());
                                println(rownum % 2, 4, "Yes");
                            }
                        } else {
                            rownum++;
                            println(0, 1, "No related admin queries identified.");
                        }

                        // RELATED CHarts
                        rownum++;
                        rownum++;

                        // State Transitions
                        println(3, 1, "Related Admin Charts:");
                        colorCell();
                        if (relatedCharts.size() > 0) {
                            rownum++;
                            println(4, 1, "Name");
                            println(4, 2, "-");
                            println(4, 3, "Description");
                            println(4, 4, "Style");
                            println(4, 5, "Is Admin");
                            colorRow(2);
                            for (WorkItem chart : relatedCharts.values()) {
                                rownum++;
                                println(rownum % 2, 1, chart.getId());
                                println(rownum % 2, 2, "");
                                println(rownum % 2, 3, chart.getField("Description").getValueAsString());
                                println(rownum % 2, 4, chart.getField("graphStyle").getValueAsString());
                                println(rownum % 2, 5, "Yes");
                            }
                        } else {
                            rownum++;
                            println(0, 1, "No related admin charts identified.");
                        }

                        // Related Reports
                        rownum++;
                        rownum++;
                        // related triggers
                        // State Transitions
                        println(3, 1, "Related Admin Reports:");
                        colorCell();
                        if (relatedReports.size() > 0) {
                            rownum++;
                            println(4, 1, "Name");
                            println(4, 2, "-");
                            println(4, 3, "Description");
                            println(4, 4, "Is Admin");
                            colorRow(2);
                            for (WorkItem query : relatedReports.values()) {
                                rownum++;
                                println(rownum % 2, 1, query.getId());
                                println(rownum % 2, 2, "");
                                println(rownum % 2, 3, query.getField("Description").getValueAsString());
                                println(rownum % 2, 4, "Yes");
                            }
                        } else {
                            rownum++;
                            println(0, 1, "No related admin reports identified.");
                        }

                        rownum++;
                    }

                } catch (APIException e) {

                    // Report the command level Exception
                    out.println("The command failed");
                    out.println(Html.logException(e));
                    typeName = "";
                }
                // }
            }
        }
        // out.println(intSession.getAbout(title));
        // Execute disconnect
        try {
            intSession.release();
        } catch (APIException ex) {
            out.println(Html.logException(ex));
        }

        sheet = workbook.getSheetAt(0);

        out.println("Sheet Name: " + sheet.getSheetName());

        ListPickListValues(intSession);
        workbook.setSheetName(0, "Pick List Values");

        // workbook.removeSheetAt(0);
        for (int sheetId = 0; sheetId < sheetNum; sheetId++) {
            for (int i = 0; i < 7; i++) {
                // removeRow(workbook.getSheetAt(sheetId), i);
            }
        }
        writeExcelFile(targetfile.replace("&1", "" + fileNum++));

    }

    public static void removeRow(XSSFSheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            XSSFRow removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    public static void addPicure(String typeName) throws IOException {

        String fileName = imageRoot + getCharsOnly(typeName) + "_Workflow.jpeg";

        File file = new File(fileName);
        if (file.exists()) {
            /* Read the input image into InputStream */
            InputStream image = new FileInputStream(fileName);
            /* Convert Image to byte array */
            byte[] bytes = IOUtils.toByteArray(image);
            /* Add Picture to workbook and get a index for the picture */
            int my_picture_id = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            /* Close Input Stream */
            image.close();
            /* Create the drawing container */
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            /* Create an anchor point */
            XSSFClientAnchor my_anchor = new XSSFClientAnchor();
            /* Define top left corner, and we can resize picture suitable from there */
            my_anchor.setCol1(8);
            my_anchor.setRow1(8);
            /* Invoke createPicture and pass the anchor point and ID */
            XSSFPicture my_picture = drawing.createPicture(my_anchor, my_picture_id);
            /* Call resize method, which resizes the image */
            my_picture.resize();
        } else {
            out.println("The file " + fileName + " does not exist!");
        }
    }

    /**
     * Look for the given type and class in the query, if found, add to the
     * reportMap
     *
     * @param queryString
     * @param itemType
     * @param docClass
     * @param report
     */
    public static void findTypeInQuery(WorkItem query, String itemType, String docClass, WorkItem mainQuery) {
        if (query != null && query.getField("queryDefinition") != null && query.getField("queryDefinition").getValueAsString() != null) {

            // ia.debug("Query name: '" + query.getName() + "'");
            // ia.debug("Query def:  '" + query.getQueryDefinition() + "'");
            String[] analysed = query.getField("queryDefinition").getValueAsString().split("\"");
            for (int j = 0; j < analysed.length; j++) {
                // out.println(j + " (qs): " + analysed[j]);
            }

            if ((analysed[0].contains("item.")) && (analysed[0].contains("item." + docClass))) {
                // add to the list if  != it contains item.segment or item.node
                relatedQueries.put(mainQuery.getId(), mainQuery);
            } else if (analysed[0].contains("subquery")) {
                // get the subquery
                analysed[0] = analysed[0].replaceAll("\\]", "\\[");
                analysed = analysed[0].split("\\[");
                for (int j = 0; j < analysed.length; j++) {
                    // out.println(j + " (sq): " + analysed[j]);
                }
                for (int j = 0; j < analysed.length; j = j + 2) {
                    if (analysed[j].contains("subquery")) {
                        findTypeInQuery(intSession.allQueries.get(analysed[j + 1]), itemType, docClass, mainQuery);
                    }
                }
            } else if (analysed.length > 3) {
                if (analysed[1].contentEquals("Type") && analysed[3].contentEquals(itemType)) {
                    // tableContentAll.add(report);        
                    relatedQueries.put(mainQuery.getId(), mainQuery);
                }

            }
        }
    }

}
