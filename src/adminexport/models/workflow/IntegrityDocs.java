package adminexport.models.workflow;

import java.io.File;

public class IntegrityDocs 
{
	public static final String iDOCS_REV = "$Revision: 1.1 $";
	private static final String os = System.getProperty("os.name");
	public static final String nl = System.getProperty("line.separator");
	public static final String fs = System.getProperty("file.separator");
	public static final File REPORT_DIR = new File(System.getProperty("user.home") + fs + "Desktop" + fs + "IntegrityDocs");
	public static final File REPORT_FILE =  new File(REPORT_DIR.getAbsolutePath() + fs + "index.htm");
	public static final File CONTENT_DIR =  new File(REPORT_DIR.getAbsolutePath() + fs + "WorkflowDocs");
	public static final File TYPES_DIR =  new File(CONTENT_DIR.getAbsolutePath() + fs + "Types");
	public static final File TRIGGERS_DIR =  new File(CONTENT_DIR.getAbsolutePath() + fs + "Triggers");
	public static final File XML_CONTENT_DIR =  new File(REPORT_DIR.getAbsolutePath() + fs + "WorkflowDocs-XML");	
	public static final File XML_VIEWSETS_DIR =  new File(XML_CONTENT_DIR.getAbsolutePath() + fs + "viewsets");
	
}
