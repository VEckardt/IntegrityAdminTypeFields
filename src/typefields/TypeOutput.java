/*
 * Copyright:      Copyright 2015 (c) Parametric Technology GmbH
 * Product:        PTC Integrity Lifecycle Manager
 * Author:         V. Eckardt, Senior Consultant ALM
 * Purpose:        Custom Developed Code
 * **************  File Version Details  **************
 * Revision:       $Revision: 1.2 $
 * Last changed:   $Date: 2017/03/17 00:04:03CET $
 */
package typefields;

import com.mks.api.response.WorkItem;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import static java.lang.System.out;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;
import static typefields.TypeFields.inputDocStream;
import static typefields.TypeFields.inputHttpStream;
import typefields.api.IntegritySession;
import static typefields.excel.ExcelWorkbook.recalcWorkbook;
import static typefields.excel.ExcelWorkbook.setValue;
import typefields.utils.Legend;

/**
 *
 * @author veckardt
 */
public class TypeOutput {

    static Map<String, WorkItem> allTypes = new LinkedHashMap<String, WorkItem>();
    // static public Map<String, WorkItem> allQueries = new LinkedHashMap<String, WorkItem>();
    static Map<String, String[]> typeFields = new HashMap<String, String[]>();
    static Map<String, String[]> stateTransMap = new LinkedHashMap<String, String[]>();
    static Map<String, String[]> sortedTypeFields = new LinkedHashMap<String, String[]>();

    static Map<String, String[]> typeEditProps = new LinkedHashMap<String, String[]>();
    // Legend
    static Legend legend = new Legend();
    static FileInputStream inputDocStream = null;
    static InputStream inputHttpStream = null;
    static XSSFWorkbook workbook;
    static String targetfile = "c:\\IntegrityAdminExport\\Integrity_TypeFields_Set&1.xlsx";
    static XSSFSheet sheet;
    static int rownum = 6;
    static int lastRowNum = -1;

    // get the xml document from a String 
    public static Document loadXMLFromString(String xml) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        InputSource is = new InputSource(new StringReader(xml));
        return builder.parse(is);
    }

    /**
     *
     * @param templateFile
     * @throws MalformedURLException
     * @throws IOException
     */
    public static void openExcelFile(String templateFile) throws MalformedURLException, IOException {

        if (templateFile.toLowerCase().startsWith("http")) {
            // URL url = new URL(URLEncoder.encode(templateFile, "UTF-8"));
            URL url = new URL(templateFile.replace(" ", "%20"));
            HttpURLConnection urlConnection = (HttpURLConnection) url.openConnection();
            inputHttpStream = urlConnection.getInputStream();
            //Access the workbook
            workbook = new XSSFWorkbook(inputHttpStream);
        } else {
            //Read the spreadsheet that needs to be updated
            inputDocStream = new FileInputStream(new File(templateFile));
            workbook = new XSSFWorkbook(inputDocStream);
        }

        //Access the worksheet, so that we can update / modify it.
        // XSSFSheet sheet = workbook.getSheet("Data"); // .getSheetAt(0);
        sheet = workbook.getSheetAt(0);
    }

    public static void writeExcelFile(String targetfile) throws IOException {
        log("Closing input streams ...", 2);
        if (inputDocStream != null) {
            inputDocStream.close();
        }
        if (inputHttpStream != null) {
            inputHttpStream.close();
        }

        // if (cntRows == 0) {
        //     log("ERROR: No file written, possibly no data selected?", 1);
        // } else {
        // delete the template rows
        //     if (mode == 1) {
        //         for (int k = endContentRow; k >= beginContentRow; k--) {
        //             deleteRow(sheet, k);
        //         }
        //     }
        // Create the output file
        FileOutputStream out = new FileOutputStream(new File(targetfile));
        {
            recalcWorkbook(workbook);
            workbook.write(out);
        }
        log("INFO: Excel file " + targetfile + " written successfully.", 1);
        // }
    }

    public static void log(String string, int i) {
        System.out.println(string);
    }

    public static void colorCell() {
//        // Aqua background
//        CellStyle style = workbook.createCellStyle();
//        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//
//        sheet.getRow(rownum).getCell(0).setCellStyle(style);
    }

    public static void colorRow(int color) {
        // Aqua background
//        CellStyle style = workbook.createCellStyle();
//        if (color == 0) {
//            style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
//        } else if (color == 1) {
//            style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
//        } else {
//            style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
//        }
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//
//        for (int i = 0; i < 30; i++) {
//            XSSFCell cell = sheet.getRow(rownum).getCell(i);
//            if (cell != null) {
//                cell.setCellStyle(style);
//            }
//        }
    }

    public static String getCharsOnly(String theString) {
        return theString.replaceAll("[^A-Za-z0-9]+", "");
    }

    public static void getFieldRow(IntegritySession intSession, List<String> fields, String ident,
            Map<String, String> copyFields, Map<String, String> significantEditFields,
            Map<String, String[]> typeProps, Map<String, String> stateTransitions,
            Map<String, String> mandatoryFields) {
        // for all fields
        for (int i = 0; i < fields.size(); i++) {
            WorkItem field = (WorkItem) intSession.allFields.get(fields.get(i));

            if (field != null) {
                // allow color setting later using #bgcolor#
                String[] data = new String[30];
                data[0] = fields.get(i);
                int col = 1;

                data[col++] = (field.getField("displayName") != null ? field.getField("displayName").getValueAsString() : "");
                // data[col++] = field.getField("description").getValueAsString();
                if (field.getField("type").getValueAsString().contentEquals("fva")) {
                    data[col++] = field.getField("description").getValueAsString() + " " + "(Relationship: "
                            + field.getField("backedBy").getValueAsString().replace(".", ", Field: ") + ")";
                } else {
                    data[col++] = field.getField("description").getValueAsString();
                }

                try {
                    String maxLength = field.getField("maxLength").getValueAsString();
                    data[col++] = field.getField("type").getValueAsString() + " (" + maxLength + ")";
                } catch (Exception e) {
                    data[col++] = field.getField("type").getValueAsString();
                }

                data[col++] = ident;
                if (copyFields.containsKey(fields.get(i))) {
                    data[col++] = "x";
                } else {
                    data[col++] = "-";
                }
                if (significantEditFields.containsKey(fields.get(i))) {
                    data[col++] = "x";
                } else {
                    data[col++] = "-";
                }

                for (Object key : typeProps.keySet()) {
                    String fieldsProp = typeProps.get(key)[0];
                    if (("," + fieldsProp + ",").contains("," + fields.get(i) + ",")) {
                        data[col++] = "x";
                    } else {
                        data[col++] = "-";
                    }
                }

                for (Object key : stateTransitions.keySet()) {
                    if (mandatoryFields.containsKey(fields.get(i) + "+" + key.toString())) {
                        data[col++] = "x";
                    } else {
                        data[col++] = "-";
                    }
                }

                typeFields.put(fields.get(i), data);
                sortedTypeFields.put(fields.get(i), data);
            }
        }
        // out.println(">>> typeFields.size() > "+typeFields.size()+"");
    }

    public static void println(int sourceRow, int col, String text) {
        XSSFRow row;
        if (rownum != lastRowNum) {

            row = sheet.createRow(rownum);
            lastRowNum = rownum;
        } else {
            row = sheet.getRow(rownum);
        }
        XSSFCell cell = row.createCell(col - 1);
        if (text != null) {
            CellStyle currentStyle = sheet.getRow(sourceRow).getCell(col - 1).getCellStyle();
            cell.setCellStyle(currentStyle);
        }
        if (text != null && !text.isEmpty()) {
            // out.println("writing to " + rownum + " - " + (col - 1));
            setValue(sheet, rownum, col - 1, text);
        }
    }
}
