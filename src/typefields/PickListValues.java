/*
 * Copyright:      Copyright 2017 (c) Parametric Technology GmbH
 * Product:        PTC Integrity Lifecycle Manager
 * Author:         Volker Eckardt, Principal Consultant ALM
 * Purpose:        Custom Developed Code
 * **************  File Version Details  **************
 * Revision:       $Revision$
 * Last changed:   $Date$
 */
package typefields;

import java.util.*;
import com.mks.api.*;
import com.mks.api.response.*;
import static java.lang.System.out;
import static typefields.TypeOutput.println;
import static typefields.TypeOutput.rownum;
import typefields.api.IntegritySession;
import typefields.utils.Html;

/**
 *
 * @author veckardt
 */
public class PickListValues {

    String title = "Pick List Values - V1.0";

    static String field = "All";

    static String[] headerFields = {"name", "displayName", "description", "default", "relevanceRule"};
    static String[] rowFieldsSC = {"label", "value", "active", "phase"};
    static String[] rowFieldsStd = {"label", "value", "active"};
    static String[] rowFields = {"label", "value", "active"};

    public static void ListPickListValues(IntegritySession intSession) throws APIException {

        intSession.allFields.clear();
        intSession.readAllFields("pick");

        Response respo;
        rownum = 7;

        Boolean showAll = (field.contentEquals(new StringBuffer("All")));
        for (String key : intSession.allFields.keySet()) {
            WorkItem wi;
            if (showAll) {
                wi = intSession.allFields.get(key);
            } else {
                wi = intSession.allFields.get(field);
            }

            String[] groupSC = {"nonmeaningful", "meaningful", "segment"};
            String[] group = {"xx"};
            if (wi.getField("name").getValueAsString().contentEquals(new StringBuffer("Shared Category"))) {
                rowFields = rowFieldsSC;
                group = groupSC;
            } else {
                rowFields = rowFieldsStd;
            }

            for (String headerField : headerFields) {
                println(5, 1, Html.capitalize(headerField));
                println(5, 2, wi.getField(headerField).getValueAsString());
                rownum++;
            }

            int pos = 1;
            // pick field values
            for (String rowField : rowFields) {
                println(5, pos++, Html.capitalize(rowField));
            }

            // im viewfield
            Command cmd = new Command(Command.IM, "viewfield");
            String fieldName = wi.getField("name").getValueAsString();
            cmd.addSelection(fieldName);
            out.println("Analysing field '" + fieldName + "' ...");

            respo = intSession.execute(cmd);
            // ResponseUtil.printResponse(respo, 1, System.out);
            WorkItemIterator wiip = respo.getWorkItems();
            while (wiip.hasNext()) {
                WorkItem wip = wiip.next();
                // get pick field values
                for (String group1 : group) {
                    Field fld = wip.getField("picks");

                    ListIterator li = fld.getList().listIterator();
                    while (li.hasNext()) {
                        Item item = (Item) li.next();
                        String phase;
                        try {
                            Field phaseField = item.getField("phase");

                            phase = (phaseField == null ? "" : item.getField("phase").getValueAsString() + "");
                        } catch (NoSuchElementException ex) {
                            phase = "";
                        }
                        // for (int i = 0; i< item.getFieldListSize(); i++) {
                        // 	item.
                        // }
                        // there is no image ... unfortunately
                        // Iterator it = item.getFields();
                        // while (it.hasNext()) {
                        // 	Field afld = (Field)it.next();
                        // 	out.println(afld.getName());
                        // }
                        if ((group.length == 1) || ((group.length > 1) && (group1.contentEquals(new StringBuffer(phase))))) {
                            pos = 1;
                            for (String rowField : rowFields) {
                                int color = rownum % 2;
                                try {
                                    println(color, pos++, item.getField(rowField).getValueAsString());
                                } catch (NoSuchElementException e) {
                                    println(color, pos++, "");
                                }
                            }
                            rownum++;
                        }

                    }
                }
            }
            rownum++;
            if (!showAll) {
                break;
            }
        }
    }
}
