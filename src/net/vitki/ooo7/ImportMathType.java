package net.vitki.ooo7;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Vector;
import java.io.InputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.BufferedInputStream;
import java.io.StringReader;
import java.io.InputStreamReader;

import org.xml.sax.InputSource;
import org.w3c.dom.*;

/**
 * @author vit
 *
 */
public class ImportMathType extends OooDocument
{
	public static final String WMF_DIR = REPLACEMENTS_DIR;
	public static final String MATH_PUB_ID = "-//OpenOffice.org//DTD Modified W3C MathML 1.01//EN";
	public static final String MATH_SYS_ID = "math.dtd";
	public static final String MATH_CONTENT_NAME = "/content.xml";
	public static final String MATH_SETTINGS_NAME = "/settings.xml";
	public static final String XHTML_PUB_ID = "-//W3C//DTD XHTML 1.1 plus MathML 2.0//EN";
	public static final String XHTML_DTD_RESOURCE = "net/vitki/ooo7/msm.dtd";

	public static final String MATHTYPE4_OLE_TYPE = "MathType 4.0 Equation";
	public static final String MATHTYPE5_OLE_TYPE = "MathType 5.0 Equation";
	public static final String MATHTYPE_TRANS_ERROR_MARK = "***TRANSLATION";
	public static final String MATH_TRANS_ERROR_ATTR = "math:mathtype-trans-error";
	public static final String OUR_TRANS_ERROR_MARK = "[*MATHTYPE_TRANSLATION_ERROR*]";

	protected HashMap formulas;
	protected Vector  mathmls;
	protected HashMap graphic_styles;
	protected HashSet foreign_objects;
	protected HashSet supported_ole_types;

	protected Document settings_template;
	protected Node setting_width;
	protected Node setting_height;

	protected int trivial_count;
	protected int trans_error_count;
	protected String formula_style_name;
	protected String formula_font_name = null;
	protected int formula_font_size = 0;
	protected boolean formula_font_italic = false;

	protected int cur_mml;

	private static class Formula {
		String     name;
		Element	   draw;
		OooFormula math;
		int        no;
		int		   width;
		int		   height;
	}

	public ImportMathType() throws Exception {
		supported_ole_types = new HashSet();
		supported_ole_types.add(MATHTYPE5_OLE_TYPE);
		formulas = new HashMap();
		mathmls = new Vector();
		foreign_objects = new HashSet();
		graphic_styles = new HashMap();
		cur_mml = 0;
		initSettingsTemplate();
		trivial_count = 0;
		trans_error_count = 0;
		formula_style_name = null;
		formula_font_name = null;
		formula_font_size = 0;
		formula_font_italic = false;
	}

    /*
     * ========================= Drawing Elements ===============================
     */
	
	protected void handleAutoStyle(String style_name, Element style_elem) {
		if (style_elem.getAttribute("style:family").equals("graphic"))
			graphic_styles.put(style_name, style_elem);
	}

	protected final Element getGraphicStyle(String name) {
		return (Element) graphic_styles.get(name);
	}

	protected void traverseContent(Element src, Element dst, int mode) throws Exception {

		if (!"draw:object-ole".equals(src.getNodeName())) {
			super.traverseContent(src, dst, mode);
			return;
		}

		// Verify this is really MathType object.
		String name = Util.normalRef(src.getAttribute("xlink:href"));
		if (!getMimeType(name).equals(MIME_TYPE_OLE)) {
			if (isDebug() && !foreign_objects.contains(name)) {
				logln("OLE ["+name+"]: mime is not OLE !");
				foreign_objects.add(name);
			}
			return;
		}
		String obj_type = getObjectType(name);
		if (!supported_ole_types.contains(obj_type)) {
			if (isDebug() && !foreign_objects.contains(name)) {
				logln("OLE ["+name+"]: the ["+obj_type+"] object is not supported !");
				foreign_objects.add(name);
			}
			return;
		}

		Element frame = (Element) src.getParentNode();
		if (!"draw:frame".equals(frame.getNodeName())) {
			if (isDebug() && !foreign_objects.contains(name)) {
				logln("OLE ["+name+"]: parent is not draw:frame (OLE type is ["+obj_type+"]) !");
				foreign_objects.add(name);
			}
			return;			
		}
		Document doc = getContentDocument();
		Element para = (Element) frame.getParentNode();
		
		logln("OLE ["+name+"]: OK parsing the ["+obj_type+"] object");
		// Corresponding WMF object is next door, we're so sure we even don't check.
		Element image = (Element) frame.getElementsByTagName("draw:image").item(0);
		String image_path = Util.normalRef(image.getAttribute("xlink:href"));
		if (getMimeType(image_path).equals("")) {
			setFileType(image_path, MIME_TYPE_WMF);
		}
		
		// Replace other manifest items
		removeFileType(name);
		setFileType(name+"/", MIME_TYPE_FORMULA);
		setFileType(name+MATH_CONTENT_NAME, "text/xml");
		setFileType(name+MATH_SETTINGS_NAME, "text/xml");
		setFileType(WMF_DIR+name, MIME_TYPE_WMF);
		
		// Correct graphic style properties.
		Element style_elem = getGraphicStyle(frame.getAttribute("draw:style-name"));
		Element graphic_props = (Element) style_elem.getElementsByTagName("style:graphic-properties").item(0);
		graphic_props.removeAttribute("draw:draw-aspect");
		graphic_props.removeAttribute("draw:visible-area-left");
		graphic_props.removeAttribute("draw:visible-area-top");
		graphic_props.removeAttribute("draw:visible-area-width");
		graphic_props.removeAttribute("draw:visible-area-height");
		
		// Replace OLE reference by formula reference
		Element new_elem = src.getOwnerDocument().createElementNS(DRAW_NS, "draw:object");
		NamedNodeMap atts = src.getAttributes();
		for (int i = 0; i < atts.getLength(); i++) {
			new_elem.setAttributeNodeNS((Attr)atts.item(i).cloneNode(false));
		}
		frame.replaceChild(new_elem, src);
				
		// Add the description
		int no = formulas.size();
		Formula formula = new Formula();
		formula.name = name;
		formula.width = convertSize(frame.getAttribute("svg:width"));
		formula.height = convertSize(frame.getAttribute("svg:height"));
		formula.draw = new_elem;
		
		// the corresponding MathML object is already in mathmls.
		formula.no = no;
		cur_mml = no + 1;
		boolean trans_error;
		String trivial;
		while(true) {
			Element mathml;
			boolean exists = (no < mathmls.size());
			if (exists) {
				mathml = (Element)mathmls.get(no);
			} else {
				if (no == mathmls.size())
					System.err.println(" error: no MathML for formula "+cur_mml+" !");
				Document d = newDocument();
				mathml = d.createElementNS(MATH_NS, "m:math");
				d.appendChild(mathml);
				Element sema = d.createElementNS(MATH_NS, "m:semantics");
				mathml.appendChild(sema);
				Element row = d.createElementNS(MATH_NS, "m:mrow");
				sema.appendChild(row);
				Element mi = d.createElementNS(MATH_NS, "m:mi");
				mi.setTextContent("NoSuchFormula"+cur_mml);
				row.appendChild(mi);	
			}
			trans_error = (mathml.getAttributeNode(MATH_TRANS_ERROR_ATTR) != null);
			if (isDebug()) {
				writeNodeToFile(makeDebugPath("a/a"+Util.num000(cur_mml)+".xml"), mathml);
				Element sem = (Element) mathml.getElementsByTagName("m:semantics").item(0);
				String text = sem.getElementsByTagName("*").item(0).getTextContent();
				log(Util.num000(cur_mml)+" = {"+text+"}");
				if (text.length() == 1)
					log(" = \\u"+Integer.toHexString(text.charAt(0)));
				if (trans_error)
					log(" (TRANS ERROR)");
			}
			Util.normalizeElement(mathml);
			formula.math = new OooFormula(mathml);
			String anno = formula.math.getAnnotation();
			String fake = formula.math.getFake();
			if (fake == null && formula_style_name != null)
				trivial = formula.math.getTrivialAnnotation();
			else
				trivial = null;
			if (isDebug()) {
				if (fake != null) {
					logln(" FAKE: "+fake);
				} else {
					log(" = \""+(trivial != null ? trivial : anno)+"\"");
					logln(trivial != null ? " (TRIVIAL)" : "");
					writeNodeToFile(makeDebugPath("e/e"+Util.num000(cur_mml)+".xml"), formula.math.getDocument());
				}
			}
			// free up memory
			if (exists)
				mathmls.set(no, null);
			mathml.getParentNode().removeChild(mathml);
			if (fake == null)
				break;
			if (exists)
				mathmls.remove(no);
		}

		if (trans_error) {
			// mark translation errors
			Node mark = doc.createTextNode(OUR_TRANS_ERROR_MARK);
			Node next = frame.getNextSibling();
			if (next == null)
				para.appendChild(mark);
			else
				para.insertBefore(mark, next);
			trans_error_count++;
		}
		
		if (trivial != null) {
			// substitute trivial formulas
			String plain = trivial;
			String subscript = null;
			String superscript = null;
			Element inline;
			int pos;
			pos = plain.indexOf('_');
			if (pos >= 0) {
				subscript = plain.substring(pos + 1);
				plain = plain.substring(0, pos);
			}
			pos = plain.indexOf('^');
			if (pos >= 0) {
				superscript = plain.substring(pos + 1);
				plain = plain.substring(0, pos);
			}
			if (plain != null) {
				inline = doc.createElementNS(TEXT_NS, "text:span");
				inline.setAttributeNS(TEXT_NS, "text:style-name", formula_style_name);
				inline.setTextContent(plain);
				para.insertBefore(inline, frame);
			}
			if (subscript != null) {
				inline = doc.createElementNS(TEXT_NS, "text:span");
				inline.setAttributeNS(TEXT_NS, "text:style-name", formula_style_name+"_subscript");
				inline.setTextContent(subscript);
				para.insertBefore(inline, frame);				
			}
			if (superscript != null) {
				inline = doc.createElementNS(TEXT_NS, "text:span");
				inline.setAttributeNS(TEXT_NS, "text:style-name", formula_style_name+"_superscript");
				inline.setTextContent(superscript);
				para.insertBefore(inline, frame);				
			}
			para.removeChild(frame);

			// mark objects as removed 
			markRemovedObject(name);
			markRemovedObject(name+"/");
			markRemovedObject(name+MATH_CONTENT_NAME);
			markRemovedObject(name+MATH_SETTINGS_NAME);
			markRemovedObject(WMF_DIR+name);
			
			formula = null; // placeholder
			trivial_count++;
		}

		formulas.put(name, formula);
		if (no % 100 == 0)
			progressPrint("["+(no/100)+"]");
	}
	
    /*
     * ========================= Formula Styles ===============================
     */
	
	private void createFormulaStyle() {
		
		if (formula_style_name == null)
			return;

		Document doc;
		Element root, all_styles, style, text_props;

		// create styles for subscripts and superscripts (we are sure they don't exist)
		doc = getContentDocument();
		root = doc.getDocumentElement();
		all_styles = (Element) root.getElementsByTagName("office:automatic-styles").item(0);

		style = doc.createElementNS(STYLE_NS, "style:style");
		style.setAttributeNS(STYLE_NS, "style:name", formula_style_name + "_subscript");
		style.setAttributeNS(STYLE_NS, "style:parent-style-name", formula_style_name);
		style.setAttributeNS(STYLE_NS, "style:family", "text");
		text_props = doc.createElementNS(STYLE_NS, "style:text-properties");
		text_props.setAttributeNS(STYLE_NS, "style:text-position", "sub 58%");
		style.appendChild(text_props);
		all_styles.appendChild(style);

		style = doc.createElementNS(STYLE_NS, "style:style");
		style.setAttributeNS(STYLE_NS, "style:name", formula_style_name + "_superscript");
		style.setAttributeNS(STYLE_NS, "style:parent-style-name", formula_style_name);
		style.setAttributeNS(STYLE_NS, "style:family", "text");
		text_props = doc.createElementNS(STYLE_NS, "style:text-properties");
		text_props.setAttributeNS(STYLE_NS, "style:text-position", "super 58%");
		style.appendChild(text_props);
		all_styles.appendChild(style);

		// find the place for the style itself (it may exist)
		doc = getStylesDocument();
		root = doc.getDocumentElement();
		all_styles = (Element) root.getElementsByTagName("office:styles").item(0);
		NodeList style_list = all_styles.getElementsByTagName("style:style");
		int size = style_list.getLength();
		int ref = -1;
		if (ref == -1) {
			for (int i = 0; i < size; i++) {
				Element e = (Element)style_list.item(i);
				if (e.getAttribute("style:family").equals("text")) {
					if (e.getAttribute("style:name").equals(formula_style_name)
						|| e.getAttribute("style:display-name").equals(formula_style_name)) {
						// such a style already exists, how nice !
						return;
					}
					ref = i;
				}
			}
		}
		if (ref == -1) {
			for (int i = 0; i < size; i++)
				if (((Element)style_list.item(i)).getAttribute("style:family").equals("paragraph"))
					ref = i;
		}

		// create the inline formula style		
		String font_decl_name = null;
		style = doc.createElementNS(STYLE_NS, "style:style");
		style.setAttributeNS(STYLE_NS, "style:name", formula_style_name);
		style.setAttributeNS(STYLE_NS, "style:display-name", formula_style_name.replaceAll("_", " "));
		style.setAttributeNS(STYLE_NS, "style:family", "text");
		text_props = doc.createElementNS(STYLE_NS, "style:text-properties");
		if (formula_font_name != null) {
			font_decl_name = formula_font_name + "_" + formula_style_name;
			text_props.setAttributeNS(STYLE_NS, "style:font-name", font_decl_name);
		}
		if (formula_font_italic) {
			String style_str = "italic";
			text_props.setAttributeNS(FO_NS, "fo:font-style", style_str);
			text_props.setAttributeNS(STYLE_NS, "style:font-style-asian", style_str);
			text_props.setAttributeNS(STYLE_NS, "style:font-style-complex", style_str);
		}
		if (formula_font_size > 0) {
			String size_str = String.valueOf(formula_font_size)+"pt";
			text_props.setAttributeNS(FO_NS, "fo:font-size", size_str);
			text_props.setAttributeNS(STYLE_NS, "style:font-size-asian", size_str);
			text_props.setAttributeNS(STYLE_NS, "style:font-size-complex", size_str);
		}
		if (true) {
			String lang_str = "none";
			text_props.setAttributeNS(FO_NS, "fo:language", lang_str);
			text_props.setAttributeNS(STYLE_NS, "style:language-asian", lang_str);
			text_props.setAttributeNS(FO_NS, "fo:country", lang_str);
			text_props.setAttributeNS(STYLE_NS, "style:country-asian", lang_str);
		}
		style.appendChild(text_props);

		// add the font
		if (font_decl_name != null) {
			Element all_fonts = (Element) root.getElementsByTagName("office:font-face-decls").item(0);
			String apos = formula_font_name.contains(" ") ? "'" : "";
			String adornments = formula_font_italic ? "Italic" : "Regular";
			Element font = doc.createElementNS(STYLE_NS, "style:font-face");
			font.setAttributeNS(STYLE_NS, "style:name", font_decl_name);
			font.setAttributeNS(STYLE_NS, "style:font-adornments", adornments);
			font.setAttributeNS(STYLE_NS, "style:font-family-generic", "roman");
			font.setAttributeNS(STYLE_NS, "style:font-pitch", "variable");
			font.setAttributeNS(SVG_NS, "svg:font-family", apos+formula_font_name+apos);
			all_fonts.appendChild(font);
		}

		// add the style
		if (ref == -1)
			all_styles.appendChild(style);
		else
			all_styles.insertBefore(style, style_list.item(ref + 1));		
	}
	
	public void setupStyler (String styler) {
		if (styler == null)
			return;
		final String def_style_name = "InlineFormulaStyle";
		final String def_font_name = "Times New Roman";
		final int def_font_size = 12;
		final boolean def_font_italic = true;
		styler = styler.trim();
		if ("-".equals(styler) || "+".equals(styler) || "default".equals(styler))
			styler = ",,,,,,";
		styler = styler.replaceAll(","," , ").replaceAll("_", " ");
		String[] tokens = styler.split(",");
		int n = tokens.length;
		if (n < 1)
			return;
		formula_style_name = tokens[0].trim();
		if ("".equals(formula_style_name))
			formula_style_name = def_style_name;
		formula_style_name = formula_style_name.replaceAll(" ", "_");
		if (n < 2)
			return;
		formula_font_name = tokens[1].trim();
		if ("".equals(formula_font_name))
			formula_font_name = def_font_name;
		if (n < 3)
			return;
		String size_str = tokens[2].trim();
		if ("".equals(size_str))
			formula_font_size = def_font_size;
		else
			formula_font_size = Integer.parseInt(size_str);
		if (n < 4)
			return;
		String style_str = tokens[3].trim();
		if (style_str.toLowerCase().startsWith("i"))
			formula_font_italic = true;
		else if ("".equals(style_str))
			formula_font_italic = def_font_italic;
		else
			formula_font_italic = false;
	}
	
    /*
     * ========================= XHTML ===============================
     */
	
    private void parseXHTML (String xht_path) throws Exception {
    	String basename = Util.getFileName(getInputPath());
		if (xht_path == null) {
			xht_path = basename + ".xhtml";
			if (!Util.fileExists(xht_path)) {
				xht_path = basename + ".xht";
				if (!Util.fileExists(xht_path)) {
					xht_path = basename + ".html";
					if (!Util.fileExists(xht_path)) {
						xht_path = basename + ".xml";
					}
				}
			}
		}
		InputStream xht_is = new FileInputStream(xht_path);
		Document xhtml;
		xhtml = getDocumentBuilder().parse(fixMSGlitches(xht_is));
		xht_is.close();
		progressPrintln("xhtml="+xht_path);
		traverseHTML(xhtml.getDocumentElement());    	
    }
    
	public void traverseHTML(Element elem) throws Exception {
		NodeList children = elem.getChildNodes();
		int n = children.getLength();
		for (int i = 0; i < n; i++) {
			Node child = children.item(i);
			if (child.getNodeType() == Node.ELEMENT_NODE) {
				if (child.getNodeName().equals("m:math")) {
					mathmls.add(child);
					Node next = child.getNextSibling();
					if (next != null && next.getNodeType() == Node.TEXT_NODE) {
						// mark translation errors
						String s = next.getNodeValue();
						if (s != null && s.trim().startsWith(MATHTYPE_TRANS_ERROR_MARK))
							((Element)child).setAttributeNS(MATH_NS, MATH_TRANS_ERROR_ATTR, "1");						
					}
				} else {
					traverseHTML((Element)child);
				}
			}
		}
	}

	// the following transformer fights XHTML produced by Microsoft Word:
	// it removes <![if...> ... <![else]> ... <![endif]>
    // and removes "--" within comments
    private InputSource fixMSGlitches(InputStream is) throws IOException {
    	
    	final int io_buf_size = 4096;
        int pos_s, pos_e, len_s, len_e;
        
    	// detect the encoding
    	String encoding = "UTF-8";
    	BufferedInputStream bis = new BufferedInputStream(is, io_buf_size);
        bis.mark(48);
        byte start_bytes[] = new byte[40];
        int start_len = bis.read(start_bytes);
        bis.reset();
        int seek_len = 0;
        while (seek_len < start_len
        		&& start_bytes[seek_len] > 31 && start_bytes[seek_len] < 127
        		&& start_bytes[seek_len] != (byte)'>')
        	seek_len++;
        String start_string = new String(start_bytes, 0, seek_len, "ISO-8859-1");
        start_bytes = null;
        final String enc_s = "encoding=\"";
        pos_s = start_string.indexOf(enc_s);
        if (pos_s > 0) {
        	len_s = enc_s.length();
        	pos_s += len_s;
        	pos_e = start_string.indexOf(enc_s, pos_s);
        	if (pos_e > 0)
        		encoding = start_string.substring(pos_s, pos_e).trim();
        }
        
        // read the stream into the buffer
        InputStreamReader isr = new InputStreamReader(bis, encoding);
        StringBuffer b = new StringBuffer(is.available());
        char cbuf[] = new char[io_buf_size];
        while(true) {
        	int len = isr.read(cbuf);
        	if (len < 0)
        		break;
        	b.append(cbuf, 0, len);
        }
        cbuf = null;
        final int blen = b.length();
        
        // substitute DTD URL by our DTD
        if (true) {
    		ClassLoader cl = this.getClass().getClassLoader();
    		String dtd_id = cl.getResource(XHTML_DTD_RESOURCE).toString();
        	pos_s = b.indexOf("<!DOCTYPE");
        	pos_s = b.indexOf("PUBLIC", pos_s) + 1;
        	for (int i = 0; i < 3; i++)
        		pos_s = b.indexOf("\"", pos_s) + 1;
        	pos_e = b.indexOf("\"", pos_s);
        	b.replace(pos_s, pos_e, dtd_id);
        }
        
        // remove "--" within comments
        if (true) {
        	final String rem_s = "<!--";
        	final String rem_e = "-->";
        	pos_e = 0;
        	len_s = rem_s.length();
        	len_e = rem_e.length();
        	while(true) {
        		pos_s = b.indexOf(rem_s, pos_e);
        		if (pos_s < 0)
        			break;
        		pos_s += len_s;
        		pos_e = b.indexOf(rem_e, pos_s);
        		if (pos_e < 0)
        			pos_e = blen;
        		while(pos_s < pos_e && b.charAt(pos_s)=='-')
        			b.setCharAt(pos_s++, ' ');
        		while(pos_e > pos_s && b.charAt(pos_e-1)=='-')
        			b.setCharAt(--pos_e, ' ');
        		for (int i = pos_s+1; i < pos_e; i++) {
        			if (b.charAt(i-1)=='-' && b.charAt(i)=='-') {
        				b.setCharAt(i-1, ' ');
        				b.setCharAt(i, ' ');
        				i++;
        			}
        		}
        	}
        }
        
        // remove non-standard MS "preprocessing" tags
        if (true) {
        	final String tag_s = "<![";
        	final String tag_e = "]>";
        	pos_e = 0;
        	len_s = tag_s.length();
        	len_e = tag_e.length();
        	while(true) {
        		pos_s = b.indexOf(tag_s, pos_e);
        		if (pos_s < 0)
        			break;
        		pos_s += len_s;
        		pos_e = b.indexOf(tag_e, pos_s);
        		pos_e = pos_e < 0 ? blen : pos_e + len_e;
        		if (!(b.charAt(pos_s)=='C' && b.charAt(pos_s+1)=='D'
        				&& b.charAt(pos_s+2)=='A' && b.charAt(pos_s+3)=='T')) {
        			// this is not <![CDATA[section]]>, but rather MS extension - remove.
        			for (int i = pos_s - len_s; i < pos_e; i++) {
        				char c = b.charAt(i);
        				if (!Character.isWhitespace(c))
        					b.setCharAt(i, ' ');
        			}
        		}
        	}
        }

        // convert to InputSource and return
        String result = b.toString();
        b.setLength(0);
        b = null;
        if (isDebug()) {
        	FileOutputStream fos = new FileOutputStream(makeDebugPath("fixed-xhtml.xml"));
        	OutputStreamWriter osw = new OutputStreamWriter(fos, encoding);
        	osw.write(result);
        	osw.close();
        }
        return new InputSource(new StringReader(result));
    }
    
    /*
     * ========================= Input / Output ===============================
     */
	
	protected boolean entryOutputHandler (String name) throws Exception {
		Formula formula = (Formula) formulas.get(name);
		if (formula == null)
			return false;
		if (formula.name != null) {
			writeDoc(name+MATH_CONTENT_NAME, formula.math.getDocument(), MATH_PUB_ID, MATH_SYS_ID);
			writeDoc(name+MATH_SETTINGS_NAME,
					createSettings(formula.width, formula.height),
					null, null);
			}
		if (cur_mml % 100 == 0)
			progressPrint("["+(cur_mml/100)+"]");
		cur_mml++;
		return true;
	}

    /*
     * ========================= Settings ===============================
     */

	public Document createSettings (int width, int height) throws Exception {
		setting_width.setNodeValue(String.valueOf(width));
		setting_height.setNodeValue(String.valueOf(height));
		return settings_template;
	}

	public void initSettingsTemplate () throws Exception {
		Document doc = settings_template = newDocument();
		Element doc_set = doc.createElementNS(OFFICE_NS, "office:document-settings");
		doc_set.setAttributeNS(OFFICE_NS, "office:version", "1.0");
		doc_set.setAttribute("xmlns:ooo", OOO_NS);
		doc_set.setAttribute("xmlns:xlink", XLINK_NS);
		doc.appendChild(doc_set);
		Element set_set = doc.createElementNS(OFFICE_NS, "office:settings");
		doc_set.appendChild(set_set);
		Element view_set = doc.createElementNS(CONFIG_NS, "config:config-item-set");
		view_set.setAttributeNS(CONFIG_NS, "config:name", "ooo:view-settings");
		set_set.appendChild(view_set);
		appendTemplateSetting(view_set, "ViewAreaTop", "int", "0");
		appendTemplateSetting(view_set, "ViewAreaLeft", "int", "0");
		setting_width = appendTemplateSetting(view_set, "ViewAreaWidth", "int", "0");
		setting_height = appendTemplateSetting(view_set, "ViewAreaHeight", "int", "0");
		Element conf_set = doc.createElementNS(CONFIG_NS, "config:config-item-set");
		conf_set.setAttributeNS(CONFIG_NS, "config:name", "ooo:configuration-settings");
		set_set.appendChild(conf_set);
		for (int i = 0; i < ooo_settings.length; i++) {
			String ooo[] = ooo_settings[i];
			appendTemplateSetting(conf_set, ooo[0], ooo[1], ooo[2]);
		}
	}

    private Node appendTemplateSetting (Element parent, String name, String type, String value)
    			throws Exception {
    	Element setting = settings_template.createElementNS(CONFIG_NS, "config:config-item");
    	setting.setAttributeNS(CONFIG_NS, "config:name", name);
    	setting.setAttributeNS(CONFIG_NS, "config:type", type);
    	Node text_node = settings_template.createTextNode(value);
    	setting.appendChild(text_node);
    	parent.appendChild(setting);
    	return text_node;
    }

    private static String ooo_settings[][] = {
        { "Alignment", "short", "1" },
        { "BaseFontHeight", "short", "12" },
        { "BottomMargin", "short", "0" },
        { "CustomFontNameFixed", "string", "Courier New" },
        { "CustomFontNameSans", "string", "Arial" },
        { "CustomFontNameSerif", "string", "Times New Roman" },
        { "FontFixedIsBold", "boolean", "false" },
        { "FontFixedIsItalic", "boolean", "false" },
        { "FontFunctionsIsBold", "boolean", "false" },
        { "FontFunctionsIsItalic", "boolean", "false" },
        { "FontNameFunctions", "string", "Times New Roman" },
        { "FontNameNumbers", "string", "Times New Roman" },
        { "FontNameText", "string", "Times New Roman" },
        { "FontNameVariables", "string", "Times New Roman" },
        { "FontNumbersIsBold", "boolean", "false" },
        { "FontNumbersIsItalic", "boolean", "false" },
        { "FontSansIsBold", "boolean", "false" },
        { "FontSansIsItalic", "boolean", "false" },
        { "FontSerifIsBold", "boolean", "false" },
        { "FontSerifIsItalic", "boolean", "false" },
        { "FontTextIsBold", "boolean", "false" },
        { "FontTextIsItalic", "boolean", "false" },
        { "FontVariablesIsBold", "boolean", "false" },
        { "FontVariablesIsItalic", "boolean", "true" },
        { "IsScaleAllBrackets", "boolean", "false" },
        { "IsTextMode", "boolean", "false" },
        { "LeftMargin", "short", "100" },
        { "LoadReadonly", "boolean", "false" },
        { "PrinterName", "string", "" },
        { "RelativeBracketDistance", "short", "5" },
        { "RelativeBracketExcessSize", "short", "5" },
        { "RelativeFontHeightFunctions", "short", "100" },
        { "RelativeFontHeightIndices", "short", "60" },
        { "RelativeFontHeightLimits", "short", "60" },
        { "RelativeFontHeightOperators", "short", "100" },
        { "RelativeFontHeightText", "short", "100" },
        { "RelativeFractionBarExcessLength", "short", "10" },
        { "RelativeFractionBarLineWeight", "short", "5" },
        { "RelativeFractionDenominatorDepth", "short", "0" },
        { "RelativeFractionNumeratorHeight", "short", "0" },
        { "RelativeIndexSubscript", "short", "20" },
        { "RelativeIndexSuperscript", "short", "20" },
        { "RelativeLineSpacing", "short", "5" },
        { "RelativeLowerLimitDistance", "short", "0" },
        { "RelativeMatrixColumnSpacing", "short", "30" },
        { "RelativeMatrixLineSpacing", "short", "3" },
        { "RelativeOperatorExcessSize", "short", "50" },
        { "RelativeOperatorSpacing", "short", "20" },
        { "RelativeRootSpacing", "short", "0" },
        { "RelativeScaleBracketExcessSize", "short", "0" },
        { "RelativeSpacing", "short", "10" },
        { "RelativeSymbolMinimumHeight", "short", "0" },
        { "RelativeSymbolPrimaryHeight", "short", "0" },
        { "RelativeUpperLimitDistance", "short", "0" },
        { "RightMargin", "short", "100" },
        { "TopMargin", "short", "0" }
    };
    
    /*
     * ============================= Other utilities ============================
     */

	protected static final int convertSize (String str) throws Exception {
		str = str.trim();
		if (str.endsWith("cm")) {
			str = str.substring(0, str.length() - 2).trim();
			return (int)(Float.parseFloat(str) * 1000);
		}
		if (str.endsWith("in")) {
			str = str.substring(0, str.length() - 2).trim();
			return (int)(Float.parseFloat(str) * 2.54 * 1000);
		}
		throw new Exception("incorrect size string ["+str+"]");
	}
	
    /*
     * ============================= main ============================
     */
	
    public static void main(String[] args) throws Exception {
    	String in_name = null;
    	String html_name = null;
    	String out_name = null;
    	String styler = null;
    	String dbg_path = null;
    	boolean bad_usage = false;
    	for (int i = 0; i < args.length; i++) {
    		if ("-x".equals(args[i])) {
    			html_name = args[++i];
    		} else if ("-o".equals(args[i])) {
    			out_name = args[++i];
    		} else if ("-d".equals(args[i])) {
    			dbg_path = args[++i];
    		} else if ("-s".equals(args[i])) {
    			styler = args[++i];
    		} else if ("-indent".equals(args[i])) {
    			INDENTATION = true;
    		} else if ("-iso".equals(args[i])) {
    			ENCODING = "ISO-8859-1";
    		} else if ("-h".equals(args[i])) {
    			SHOW_PROGRESS = true;
    		} else if (args[i].startsWith("-")) {
    			bad_usage = true;
    			break;
    		} else if (in_name != null){
    			bad_usage = true;
    			break;
    		} else {
    			in_name = args[i];
    		}
    	}
		if (bad_usage || in_name == null) {
			System.err.println("usage: java -jar ImportMathType.jar"
								+" [-d debug_dir|+] [-iso] [-indent] [-h]"
								+" [-s Style?,Font?,Size?,Italic?|+]"
								+" [-x <in.xhtml>] [-o <out.odt>] <in.odt> ");
			System.exit(1);
		}
		progressPrintln("setting up...");
		ImportMathType w2o = new ImportMathType();
		w2o.run(in_name, html_name, out_name, dbg_path, styler);
	}
    
	public void run (String odt_path, String xht_path, String out_path,
						String dbg_path, String styler) throws Exception {
		setInputPath(odt_path);
		setOutputPath(out_path);
		setupLogging(dbg_path);
		if (isDebug()) {
			Util.cleanDirectory (makeDebugPath("a"));			
			Util.cleanDirectory (makeDebugPath("e"));			
		}
		setupStyler(styler);
		System.err.println("start: "+odt_path);
		progressPrint("parsing... ");
		parseInputDocument();
		createFormulaStyle();
		parseXHTML(xht_path);
		traverseContent();
		progressPrintln(" OK");
		if (trivial_count > 0)
			System.err.println("info: substituted "+trivial_count+" trivial formulas "
								+"out of "+formulas.size()+" total.");
		int delta = mathmls.size() - formulas.size();
		if (delta > 0) {
			// if there are excess formulas in XHTML, check that they are all fake
			for (int i = formulas.size(); i < mathmls.size(); i++) {
				Element mathml = (Element)mathmls.get(i);
				Element sem = (Element) mathml.getElementsByTagName("m:semantics").item(0);
				String text = sem.getElementsByTagName("*").item(0).getTextContent();
				Util.normalizeElement(mathml);
				OooFormula math = new OooFormula(mathml); // drop the result
				String fake = math.getFake();
				log(Util.num000(i)+"(EXCESS) = {"+text+"} = \""+math.getAnnotation()+"\" ");
				if (text.length() == 1)
					log("= '\\u"+Integer.toHexString(text.charAt(0))+"' ");
				logln(fake != null ? "- OKAY, is a FAKE("+fake+")" : "NOT FAKE !!!");
				if (fake != null)
					delta--;
			}
		}
		if (delta > 0)
			System.err.println("warning: "+delta+" excess formulas in XHTML !");
		else if (delta < 0)
			System.err.println("warning: could not find "+delta+" formulas in XHTML !");
		if (trans_error_count != 0)
			System.err.println("warning: "+trans_error_count
								+ " MathType translation errors are marked as "
								+ OUR_TRANS_ERROR_MARK);
		progressPrint("dumping... ");
		cur_mml = 0;
		dumpOutput();
		progressPrintln(" OK");
		closeLog();
		System.err.println("done: "+getOutputPath());
	}
}

