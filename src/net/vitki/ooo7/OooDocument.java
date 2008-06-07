package net.vitki.ooo7;

import java.util.zip.ZipFile;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.TreeSet;
import java.text.Collator;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.BufferedReader;
import java.io.StringReader;
import java.io.InputStreamReader;
import java.io.PrintStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.dom.DOMSource;  
import javax.xml.transform.stream.StreamResult;

import org.xml.sax.InputSource;
import org.w3c.dom.*;

/**
 * @author vit
 *
 */
public class OooDocument
{
	public static final String MANIFEST_ENTRY = "META-INF/manifest.xml";
	public static final String MANIFEST_PUB_ID = "-//OpenOffice.org//DTD Manifest 1.0//EN";
	public static final String MANIFEST_SYS_ID = "Manifest.dtd";
	public static final String MANIFEST_TAG_LIST = "manifest:manifest";
	public static final String MANIFEST_TAG_FILE = "manifest:file-entry";
	public static final String MANIFEST_TAG_PATH = "manifest:full-path";
	public static final String MANIFEST_TAG_TYPE = "manifest:media-type";
	public static final String CONTENT_ENTRY = "content.xml";
	public static final String STYLES_ENTRY = "styles.xml";
	public static final String REPLACEMENTS_DIR = "ObjectReplacements/";

	public static final String MIME_TYPE_OLE = "application/vnd.sun.star.oleobject";
	public static final String MIME_TYPE_FORMULA = "application/vnd.oasis.opendocument.formula";
	public static final String MIME_TYPE_WMF = "application/x-openoffice-gdimetafile;windows_formatname=\"GDIMetaFile\"";

	public static final String CHART_NS = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
	public static final String CONFIG_NS = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
	public static final String DC_NS = "http://purl.org/dc/elements/1.1/";
	public static final String DOM_NS = "http://www.w3.org/2001/xml-events";
	public static final String DR3D_NS = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
	public static final String DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
	public static final String FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
	public static final String FORM_NS = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
	public static final String MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
	public static final String MATH_NS = "http://www.w3.org/1998/Math/MathML";
	public static final String META_NS = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
	public static final String NUMBER_NS = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
	public static final String OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
	public static final String OOO_NS = "http://openoffice.org/2004/office";
	public static final String OOOC_NS = "http://openoffice.org/2004/calc";
	public static final String OOOW_NS = "http://openoffice.org/2004/writer";
	public static final String SCRIPT_NS = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
	public static final String STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
	public static final String SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
	public static final String TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
	public static final String TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
	public static final String XFORMS_NS = "http://www.w3.org/2002/xforms";
	public static final String XLINK_NS = "http://www.w3.org/1999/xlink";	
	public static final String XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";
	public static final String XSD_NS = "http://www.w3.org/2001/XMLSchema";

	public static final String BEENHERE_ATTR = "marked-as-been-here";
	public static final String BEENHERE_TEXT = "\uffff\uffff*marked-as-been-here*\uffff\uffff";
	public static final int BEENHERE_FLAG = 0x8000;

	public static String ENCODING = "UTF-8";
	public static boolean INDENTATION = false;
	public static boolean SHOW_PROGRESS = false;

	private String in_path;
	private String out_path;
	private ZipFile in_zip;
	private ZipOutputStream out_zip;

	private Document content;
	private Document styles;
	private DocumentBuilder dbuilder;

	private HashMap file_types;
	private HashSet removed_files;

	private boolean debug;
	private String dbg_dir;	
	private PrintStream log_ps;

	public OooDocument() throws Exception {
		file_types = new HashMap();
		removed_files = new HashSet();
		debug = false;
		dbg_dir = null;	
		log_ps = null;
		// carefully choose document builder
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setNamespaceAware(true);
		dbf.setValidating(false);
		dbf.setCoalescing(true);
		dbf.setIgnoringElementContentWhitespace(true);
		dbf.setIgnoringComments(true);
		dbuilder = dbf.newDocumentBuilder();
	}

	protected Document newDocument() {
		return dbuilder.newDocument();
	}

	protected DocumentBuilder getDocumentBuilder() {
		return dbuilder;
	}

	// carefully choose serializer/transformer

	public static boolean USE_SAXON = false;
	public static boolean CHECK_FOR_SAXON = true;

	static {
		final String saxon_tf = "net.sf.saxon.TransformerFactoryImpl";
		if (CHECK_FOR_SAXON) {
			try {
				if (ClassLoader.getSystemClassLoader().loadClass(saxon_tf) != null)
					USE_SAXON = true;
			} catch (Exception e) {
				USE_SAXON = false;
			}
		}
		if (USE_SAXON)
			System.setProperty("javax.xml.transform.TransformerFactory", saxon_tf);
	}

    /*
     * ========================= Traverser ===============================
     */

	protected void traverseContent() throws Exception {
		Element root_el = getContentDocument().getDocumentElement();
		Element styles_el = getElementByTag(root_el, "office:automatic-styles");
		NodeList style_list = styles_el.getElementsByTagName("style:style");
		int list_len = style_list.getLength();
		for (int i = 0; i < list_len; i++) {
			Element style_elem = (Element)style_list.item(i);
			String style_name = style_elem.getAttribute("style:name");
			if ("style:style".equals(style_elem.getNodeName()) && !"".equals(style_name))
				handleAutoStyle(style_name, style_elem);
		}
		Element text_el = getElementByTag(getElementByTag(root_el, "office:body"), "office:text");
		traverseContent(text_el, 0);		
	}

	protected void traverseContent(Element src, int mode) throws Exception {
		traverseContent(src, null, mode);
	}

	protected void traverseContent(Node node, Element dst, int mode) throws Exception {
		switch(node.getNodeType()) {
		case Node.ELEMENT_NODE:
			Element elem = (Element)node;
			if (!"1".equals(elem.getAttribute(BEENHERE_ATTR)))
				traverseContent(elem, dst, mode);
			else if ((mode & BEENHERE_FLAG) != 0)
				elem.setAttribute(BEENHERE_ATTR, "1");
			break;
		case Node.TEXT_NODE:
			String text = node.getNodeValue();
			if (!BEENHERE_TEXT.equals(text))
				traverseContent(text, dst, mode);
			else if ((mode & BEENHERE_FLAG) != 0)
				node.setNodeValue(BEENHERE_TEXT);
			break;
		}
	}

	protected void traverseContent(Element elem, Element dst, int mode) throws Exception {
		NodeList nl = elem.getChildNodes();
		for (int i = 0; i < nl.getLength(); i++)
			traverseContent(nl.item(i), dst, mode);
	}

	protected void traverseContent(String text, Element dst, int mode)  { }

	protected void handleAutoStyle(String style_name, Element style_elem) {	}

	protected static final Element getElementByTag(Node in, String tag) {
		if (in != null && in.getNodeType() == Node.ELEMENT_NODE) {
			NodeList nl = ((Element)in).getElementsByTagName(tag);
			if (nl != null && nl.getLength() > 0)
				return (Element) nl.item(0);
		}
		return null;
	}

	protected final void beenHere (Node node) {
		if (node.getNodeType() == Node.ELEMENT_NODE)
			((Element)node).setAttribute(BEENHERE_ATTR, "1");
		else
			node.setNodeValue(BEENHERE_TEXT);
	}

    /*
     * ========================= Inline Objects ===============================
     */

	protected void markRemovedObject(String name) {
		file_types.remove(name);
		removed_files.add(name);
	}

	public static final int magic_range_start = 0x800;
	public static final int magic_range_end = 0x900;
	public static final int magic_min_len = 4;

	protected String getObjectType(String name) throws IOException {
		StringBuffer sb = new StringBuffer();
		InputStream is = in_zip.getInputStream(in_zip.getEntry(name));
		byte[] bb = new byte[magic_range_start];
		int n = is.read(bb);
		if (n < bb.length) {
			is.close();
			return "<TOO_SMALL_OLE>";
		}
		bb = new byte[magic_range_end - magic_range_start];
		n = is.read(bb);
		is.close();
		for (int i = 0; i < n; i++) {
			byte b = bb[i];
			if (b > 31 && b < 127)
				sb.append((char)b);
			else if (b == 0 && sb.length() >= magic_min_len)
				break;
			else
				sb.setLength(0);
		}
		return sb.length() >= magic_min_len ? sb.toString().trim() : "<MAGIC_NOT_FOUND>";
	}

    /*
     * ========================= Input / Output ===============================
     */

	protected void setInputPath (String path) {
		in_path = Util.getFileName(path) + ".odt";
	}

	protected String getInputPath () {
		return in_path;
	}

	protected void setOutputPath (String path) {
		out_path = path != null ? path : Util.getFileName(in_path) + "_out.odt";
	}

	protected String getOutputPath () {
		return out_path;
	}

	protected InputStream getEntryStream (String name) throws IOException {
		return in_zip.getInputStream(in_zip.getEntry(name));
	}

	protected void parseInputDocument () throws Exception {
		in_zip = new ZipFile(new File(in_path));
		InputStream is;
		is = getEntryStream(CONTENT_ENTRY);
		content = dbuilder.parse(is);
		is.close();
		is = getEntryStream(STYLES_ENTRY);
		styles = dbuilder.parse(is);
		is.close();
		parseManifest();
	}
	
	protected Document getContentDocument() {
		return content;
	}

	protected Document getStylesDocument() {
		return styles;
	}

	protected boolean entryOutputHandler (String name) throws Exception {
		return false;
	}

	protected void dumpOutput() throws Exception {
		out_zip = new ZipOutputStream(new FileOutputStream(out_path));
		writeManifest();
		writeDoc(CONTENT_ENTRY, content, null, null);
		writeDoc(STYLES_ENTRY, styles, null, null);
		//content = null; // free up memory
		Enumeration enu = in_zip.entries();
		byte[] tmp_buf = new byte[32768];
		while (enu.hasMoreElements()) {		
			ZipEntry ze = (ZipEntry) enu.nextElement();
			String name = Util.normalRef(ze.getName());
		    if (name.equals(CONTENT_ENTRY) || name.equals(MANIFEST_ENTRY)
				|| name.equals(STYLES_ENTRY)) {
		        // already dumped
		    	continue;
		    }
			if (removed_files.contains(name)) {
				// removed
				continue;
			}
			if (!entryOutputHandler(name)) {
				out_zip.putNextEntry(ze);
				InputStream is = getEntryStream(name);
				while(true) {
					int num = is.read(tmp_buf);
					if (num < 0)
						break;
					out_zip.write(tmp_buf, 0, num);
				}
				is.close();
				out_zip.closeEntry();
			}
		}
		out_zip.finish();
		out_zip.close();
		out_zip = null;
	}

	protected void writeDoc(OutputStream os, Document doc, String pub_id, String sys_id) throws Exception {
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer ddumper = tf.newTransformer();
		ddumper.setOutputProperty(OutputKeys.INDENT, INDENTATION ? "yes" : "no");
		ddumper.setOutputProperty(OutputKeys.ENCODING, ENCODING);
		if (pub_id != null)
			ddumper.setOutputProperty(OutputKeys.DOCTYPE_PUBLIC, pub_id);
		if (sys_id != null)
			ddumper.setOutputProperty(OutputKeys.DOCTYPE_SYSTEM, sys_id);
		ddumper.transform(new DOMSource(doc), new StreamResult(os));
	}

	protected void writeDoc(String name, Document doc, String pub_id, String sys_id) throws Exception {
		ZipEntry entry = new ZipEntry(name);
		entry.setMethod(ZipEntry.DEFLATED);
		out_zip.putNextEntry(entry);
		writeDoc(out_zip, doc, pub_id, sys_id);
		out_zip.closeEntry();
	}

	protected void writeNodeToFile(String name, Node node) throws Exception {
		Document doc;
		switch(node.getNodeType()) {
		case Node.DOCUMENT_NODE:
			doc = (Document)node;
			break;
		case Node.ELEMENT_NODE:
			doc = dbuilder.newDocument();
			doc.appendChild(doc.importNode(node, true));
			break;
		default:
			throw new Exception("internal error in writeNodeToFile()");
		}
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer ddumper = tf.newTransformer();
		ddumper.setOutputProperty(OutputKeys.INDENT, "yes");
		ddumper.setOutputProperty(OutputKeys.ENCODING, ENCODING);
		FileOutputStream fos = new FileOutputStream(name);
		ddumper.transform(new DOMSource(doc), new StreamResult(fos));
		fos.close();
	}


	// the following hack is borrowed from the Writer2Latex project.
	// it removes DTD from the document. 
    protected static InputSource withoutDtd(InputStream is) throws IOException {
        BufferedReader br = new BufferedReader(new InputStreamReader(is, "UTF-8"));
        StringBuffer buffer = new StringBuffer(is.available());
        String str = null;
        while ((str = br.readLine()) != null) {
            int sIndex = str.indexOf("<!DOCTYPE");
            if (sIndex > -1) {
                buffer.append(str.substring(0, sIndex));
                int eIndex = str.indexOf('>', sIndex + 8 );
                if (eIndex > -1) {
                    buffer.append(str.substring(eIndex + 1, str.length()));
                    // FIX (HJ): Preserve the newline
                    buffer.append("\n");
                } else {
                    // FIX (HJ): More than one line. Search for '>' in following lines
                    boolean bOK = false;
                    while ((str = br.readLine())!=null) {
                        eIndex = str.indexOf('>');
                        if (eIndex>-1) {
                            buffer.append(str.substring(eIndex+1));
                            // FIX (HJ): Preserve the newline
                            buffer.append("\n");
                            bOK = true;
							break;
                        }
                    }
                    if (!bOK) { throw new IOException("Invalid XML"); }
                }
            } else {
                buffer.append(str);
                // FIX (HJ): Preserve the newline
                buffer.append("\n");
            }
        }
        StringReader r = new StringReader(buffer.toString());
        return new InputSource(r);
    }    
    
    /*
     * ========================= Manifest ===============================
     */

	protected void parseManifest () throws Exception {
		InputStream is = getEntryStream(MANIFEST_ENTRY);
		Document manifest = dbuilder.parse(withoutDtd(is));
		is.close();
		NodeList file_list = manifest.getDocumentElement().getElementsByTagName("*");
		for (int i = 0; i < file_list.getLength(); i++) {
			Element elem = (Element)file_list.item(i);
			String path = Util.normalRef(elem.getAttribute(MANIFEST_TAG_PATH));
			String mime = elem.getAttribute(MANIFEST_TAG_TYPE);
			file_types.put(path, mime);
			//System.out.println("mime: "+path+" => "+mime);
		}
	}

	protected void writeManifest() throws Exception {
		Document doc = dbuilder.newDocument();
		Element list = doc.createElementNS(MANIFEST_NS, MANIFEST_TAG_LIST);
		doc.appendChild(list);
		TreeSet sorted_paths = new TreeSet(Collator.getInstance());
		sorted_paths.addAll(file_types.keySet());
		Iterator keys = sorted_paths.iterator();
		while(keys.hasNext()) {
			String path = (String) keys.next();
			Element file = doc.createElementNS(MANIFEST_NS, MANIFEST_TAG_FILE);
			file.setAttributeNS(MANIFEST_NS, MANIFEST_TAG_PATH, path);
			file.setAttributeNS(MANIFEST_NS, MANIFEST_TAG_TYPE, getMimeType(path));
			list.appendChild(file);
		}
		writeDoc(MANIFEST_ENTRY, doc, MANIFEST_PUB_ID, MANIFEST_SYS_ID);
	}

	protected void setFileType (String path, String type) {
		file_types.put(path, type);		
	}

	protected void removeFileType (String path) {
		file_types.remove(path);		
	}

    /*
     * ============================= Other utilities ============================
     */

	protected final String getMimeType (String path) {
		return (String)file_types.get(path);
	}

    static final void progressPrint (String str) {
    	if (SHOW_PROGRESS)
    		System.err.print(str);
    }

    static final void progressPrintln (String str) {
    	if (SHOW_PROGRESS)
    		System.err.println(str);
    }

    /*
     * ============================= Debugging ============================
     */

	protected void setupLogging (String dbg_path) throws IOException {
		if (dbg_path == null)
			return;
		dbg_path = dbg_path.trim();
		if ("-".equals(dbg_path) || "+".equals(dbg_path) || "same".equals(dbg_path)) {
			dbg_path = (new File(out_path)).getAbsolutePath();
			int pos = dbg_path.lastIndexOf(File.separatorChar);
			dbg_path = pos < 0 ? "." : dbg_path.substring(0, pos);
		}
		File dir = new File(dbg_path);
		if (!dir.isDirectory())
			throw new IOException("debugging directory "+dbg_path+"does not exist");
		dbg_dir = dir.getAbsolutePath() + "/";
		debug = true;
		String purename = Util.getFileName(in_path);
		int pos = purename.lastIndexOf(File.separatorChar);
		if (pos >= 0)
			purename = purename.substring(pos+1);
		log_ps = new PrintStream(new FileOutputStream(dbg_dir+"dbg-"+purename+".log"));			
	}

	protected void closeLog () {
		if (debug)
			log_ps.close();
		log_ps = null;
	}

	protected String makeDebugPath (String name) {
		return dbg_dir + name;
	}

	protected boolean isDebug() {
		return debug;
	}

	protected void log(String s) {
		if (debug)
			log_ps.print(s);
	}

	protected void logln(String s) {
		if (debug)
			log_ps.println(s);
	}

}

