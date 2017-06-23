package typefields.utils;

import java.util.Date;
import java.util.Map;

import com.mks.api.response.APIException;
import com.mks.api.response.WorkItem;

public class Html {

  // this section is here just to make it work somehow
    // will be enhanced soon
    public static String[] titleReports = {"Type Fields", "Pick List Values", "Recent Changes",
        "Property Checker", "Project Checkpoint", "Currently Unused Fields", "Type Usage", ""};
    public static String[] descrReports = {
        "This report displays all settings related to an Integrity type, including states, properties, related triggers",
        "This report displays all pick list values for the selected fields",
        "This report lists all recently changed objects in the Integrity Administrator.",
        "The Property Checker validates if the type properties are set correctly.",
        "List of all checkpoints for the selected project. This report is here just for demo purposes.",
        "This reports lists all unused fields that you may reuse later. Remember that such fields are possibly part of the Type history.",
        "This report summarizes the items created by type and informs you about the last modified Item date. <br>You can use the result to determine which types you may hide to the users. This report might run a moment!",
        "Project Test State"};
    public static String[] jspReports = {"TypeFields.jsp", "PickListValues.jsp",
        "RecentChanges.jsp", "PropertyChecker.jsp", "ProjectCheckpoints.jsp", "UnusedFields.jsp",
        "TypeUsage.jsp", "TestState.jsp"};
    public static String[] typeReports = {"report", "report", "tool", "tool", "Source Report",
        "report", "tool", "report"};

    // convert to html
    public static String toHtml(String text) {
        return text.replaceAll("<", "&lt;").replaceAll(">", "&gt;");
    }

    // get td tags
    public static final String td(String text) {
        return ("<td>" + text + "</td>");
    }

    // get th tags
    public static final String th(String text) {
        return ("<th>" + text + "</th>");
    }

    // get h3 tags
    public static final String h3(String text) {
        return ("<h3>" + text + "</h3>");
    }

    // convert to uppercase
    public static final String capitalize(String text) {
        return Character.toUpperCase(text.charAt(0)) + text.substring(1);
    }

    // get title and description
    public static final String getTitle(String text, int index) {
        Date date = new java.util.Date();
        return "<div id=\"index\"><a href=\"Index.jsp\">Index</a></div><div align=right>"
                + date.toString() + "</div>" + "<h1>" + text + "</h1><div style='text-align:center;'>"
                + Html.descrReports[index - 1] + "</div>";
    }

  // returns the exception
    // has to be improved!
    public static String logException(APIException ae) {
        return ("<br>Exception logging " + "<br> message:     " + ae.getMessage()
                + "<br> class:       " + ae.getClass().getName() + "<br> exceptionId: " + ae
                .getExceptionId());
    }

    // generates a well looking header entry - short form
    public static String getHeaderField(WorkItem wi, String fieldName) {
        return getHeaderField(wi, fieldName, "");
    }

    // generates a well looking header entry
    public static String getHeaderField(WorkItem wi, String fieldName, String addText) {
        String value = wi.getField(fieldName).getValueAsString();
        String data = "";
        if (value != null && !value.equals("null")) {
            String fName = Html.capitalize(fieldName).replaceAll("(\\p{Ll})(\\p{Lu})", "$1 $2");

            data = ("" + fName + ":");
            if (fName.contentEquals("Name")) {
                data = data + Html.td("<b>" + value + "&nbsp;" + addText + "</b>") + "</tr>";
            } else {
                data = data + Html.td(value + "&nbsp;" + addText) + "</tr>";
            }
        }
        return data;
    }

    // gets a standartized item select box
    public static String itemSelectField(String typeName, Map objects, Boolean addAll,
            String selected1, String selected2) {
        String data = "<hr><table border=1>";
        // BEGIN: Build the selection list for types
        data = data + ("<tr>");
        data = data + ("<form name=form1>");

        data = data + getOptionBox(typeName, objects, addAll, selected1, selected2);

        data = data + (Html.td("<input type=\"submit\" value=\"Refresh\">"));
        data = data + ("</form>");
        data = data + ("</tr>");
        data = data + ("</table>");
        return data;
    }

    public static String getOptionBox(String typeName, Map objects, Boolean addAll,
            String selected1, String selected2) {
        String data = (Html.td("Display details for " + typeName));
        data = data + ("<td>");
        data = data + ("<select name=" + typeName + " onChange='this.form.submit()'>");

        if (addAll) {
            if (selected1 != null && !selected1.equals("") && selected1.equals("All")) {
                data = data + ("<option selected value='All'>* All *");
            } else {
                data = data + ("<option value='All'>* All *");
            }
        }

        for (Object key : objects.keySet()) {
            if (selected1 != null && !selected1.equals("") && selected1.equals(key.toString())) {
                data = data + ("<option selected value='" + key.toString() + "'>" + key.toString());
            } else {
                data = data + ("<option value='" + key.toString() + "'>" + key.toString());
            }
        }

        data = data + ("</select>");
        data = data + ("</td>");

        if (selected2 != null) {

            data = data + (Html.td("display") + "<td>");
            data = data + ("<select name=displayMode onChange='this.form.submit()'>");

            String[] displayModes = new String[]{"All Rows", "Issues Only"};

            for (int d = 0; d < displayModes.length; d++) {
                if (selected2 != null && !selected2.equals("") && selected2.equals(displayModes[d])) {
                    data = data + ("<option selected value='" + displayModes[d] + "'>" + displayModes[d]);
                } else {
                    data = data + ("<option value='" + displayModes[d] + "'>" + displayModes[d]);
                }
            }

            data = data + ("</select>");
            data = data + ("</td>");
        }
        return data;
    }

    // produces an index section
    public static String getIndexSection(String sectionName, String typeName) {
        String data = "<hr><h3>" + sectionName
                + "</h3><table border=1><tr><th style=\"width: 175px;\">" + capitalize(typeName)
                + "</th><th>Description</th></tr>";
        for (int i = 0; i < Html.typeReports.length; i++) {
            if (Html.typeReports[i].contentEquals(typeName)) {
                data = data + ("<tr>");
                data = data
                        + ("<td><a href=\"" + Html.jspReports[i] + "\">" + Html.titleReports[i] + "</a></td>");
                data = data + ("<td>" + Html.descrReports[i] + "</td>");
                data = data + ("</tr>");
            }
        }
        data = data + "</table>";
        return data;
    }
}
