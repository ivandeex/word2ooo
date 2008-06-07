package net.vitki.ooo7;

import java.util.Vector;
import java.util.Hashtable;
import java.util.HashMap;
import java.util.Iterator;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.StringReader;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.*;
import org.xml.sax.InputSource;

import writer2latex.latex.style.I18n;

/**
 * @author vit
 */
public class DocBookToMediaWiki
{
	public static int SPLIT_SIZE = 8192;
	Document content;
	Chunk out;
	boolean pre_mode;
	boolean no_newlines;
	boolean in_head;
	String in_dir;
	String out_dir;
	String page_link;
	String contents_link;
	String prev_link;
	String up_link;
	StringBuffer next_link_buf;
	int next_link_pos;
	StarMathConverter smc;
	StringBuffer contents_buf;
	HashMap files;
	int def_title_no;

	public DocBookToMediaWiki() {
		pre_mode = no_newlines = in_head = false;
		I18n i18n = new MyIn();
		smc = new StarMathConverter(i18n, new Hashtable(), true);
		def_title_no = 0;
	}

	public static class MyIn extends I18n {

		public MyIn() {
			super("ascii");
		}

		public String convert (String s, boolean bMathMode, String sLang) {
			// MediaWiki's TEX is nervous about non-latin characters.
			// Let's help it.
			int n = s.length();
			StringBuffer b = new StringBuffer(n+n);
			boolean changed = false;
			for (int i = 0; i < n; i++) {
				char c = s.charAt(i);
				if (c >= 0x410 && c < 0x450) {
					b.append(translit[c-0x410]);
					changed = true;
				} else {
					b.append(c);
				}
			}
			return changed ? b.toString() : s;
		}

		static final String translit[] = {
			"A","B","V","G","D","E","ZH","Z","I","J","K","L","M","N","O","P","R",
			"S","T","U","F","H","C","CH","SH","SCH","'","Y","'","AE","YU","YA",
			"a","b","v","g","d","e","zh","z","i","j","k","l","m","n","o","p","r",
			"s","t","u","f","h","c","ch","sh","sch","'","y","'","ae","yu","ya"
		};
	}

	/*
     * ===================== Main ===========================
     */

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

    public void run (String in_path, String new_out_dir, String page_name) throws Exception {
    	String basename = Util.getFileName(in_path);
    	out_dir = new_out_dir;
    	if (out_dir == null)
    		out_dir = basename + ".wiki";    	
    	int pos = basename.lastIndexOf('/');
    	if (pos >= 0)
    		basename = basename.substring(pos+1);
    	if (page_name == null)
    		page_name = basename;
    	File in_file = new File(in_path);
    	in_dir = in_file.getAbsoluteFile().getParent();
    	Util.cleanDirectory(out_dir);
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		dbf.setNamespaceAware(true);
		dbf.setValidating(false);
		dbf.setCoalescing(true);
		dbf.setIgnoringElementContentWhitespace(true);
		dbf.setIgnoringComments(true);
		DocumentBuilder dbuilder = dbf.newDocumentBuilder();
		FileInputStream fis = new FileInputStream(in_path);
		content = dbuilder.parse(withoutDtd(fis));
		fis.close();
		Util.normalizeElement(content.getDocumentElement());
		contents_buf = new StringBuffer();
		out = new Chunk("", 0);
		traverseElement(content.getDocumentElement());
		page_link = page_name;
		contents_link = page_name + "/Contents";
		up_link = prev_link = null;
		next_link_buf = null;
		next_link_pos = -1;
		files = new HashMap();
		outputSection(out, "out", page_name, null, 0);
		Iterator file_iter = files.keySet().iterator();
		while (file_iter.hasNext()) {
			String path = (String) file_iter.next();
			StringBuffer buf = (StringBuffer) files.get(path);
    		FileOutputStream fos = new FileOutputStream(path);
    		OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
		    osw.write(buf.toString());
		    osw.close();
		}
		files.clear();
		String file_path = out_dir + "/contents.txt";
		FileOutputStream fos = new FileOutputStream(file_path);
		OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
		osw.write(contents_buf.toString());
		osw.close();
		System.out.println("done "+out_dir);
    }

    public static void main (String[] args) throws Exception {
    	String in_path = null;
    	String out_path = null;
    	String page_name = null;
    	boolean bad_usage = false;
    	for (int i = 0; i < args.length; i++) {
    		String opt = args[i];
    		if ("-p".equals(opt) || "-pagename".equals(opt)) {
    			page_name = args[++i];
    		} else if ("-i".equals(opt) || "-input".equals(opt)) {
    			in_path = args[++i];
    		} else if ("-o".equals(opt) || "-outdir".equals(opt)) {
    			out_path = args[++i];
    		} else if ("-s".equals(opt) || "-splitsize".equals(opt)) {
    			SPLIT_SIZE = Integer.parseInt(args[++i]);
    		} else if ("-s-".equals(opt) || "-nosplit".equals(opt)) {
    			SPLIT_SIZE = -1;
    		} else if (args[i].startsWith("-")) {
    			bad_usage = true;
    			break;
    		} else if (in_path != null){
    			bad_usage = true;
    			break;
    		} else {
    			in_path = args[i];
    		}
    	}
		if (bad_usage || in_path == null) {
			System.err.println("usage: java DocBookToMediaWiki"
								+" [-outdir <output dir>] [-pagename <page name>]"
								+" [-splitsize <bytes>] [-nosplit]"
								+" [-input] <input file>");
			System.exit(1);
		}
    	DocBookToMediaWiki db2mw = new DocBookToMediaWiki();
    	db2mw.run(in_path, out_path, page_name);
    }

	/*
     * ===================== Serialization ===========================
     */

	public void start(String text) {
		tag("<"+text+">");
	}

	public void end(String text) {
		tag("</"+text+">");
	}

	public void empty(String text) {
		tag("<"+text+"/>");
	}

	public void tag(String text) {
		out.append(text);
	}

	public void add(String text) {
		int n = text.length();
		char c = '?', p;
		for (int i = 0; i < n; i++) {
			p = c;
			c = text.charAt(i);
			if (no_newlines && c == '\n')
				c = ' ';
			if (!pre_mode && Character.isWhitespace(p) && Character.isWhitespace(c))
				continue;
			if (pre_mode && c == '\n') {
				out.append("\n  ");
				continue;
			}
			switch(c) {
			case '<': out.append("&lt;"); continue;
			case '>': out.append("&gt;"); continue;
			case '&': out.append("&amp;"); continue;
			case 0xa0: out.append("&nbsp;"); continue;
			case 0xab: out.append('"'); continue; // <<
			case 0xbb: out.append('"'); continue; // >>
			case 0xad: continue; // hyphenation
			case 0x2013: out.append('-'); continue; // dash
			}
			if (c > 127 && (c < 0x400 || c > 0x450))
				out.append("&#x"+Integer.toHexString(c)+";");
			else
				out.append(c);
		}
	}

    public void nl() {
    	out.append("\n");
    }

	/*
     * ===================== Sections / Output ===========================
     */

    public void makeSection(Element elem, String title, int level) throws Exception {
    	title = Util.normalizeText(title);
    	if (Util.isContentsTitle(title))
    		return;
    	Chunk parent = out;
    	out = new Chunk(title, level);
    	parent.addSub(out);
    	traverseChildren(elem);
    	out = parent;
    }

    public void outputSection(Chunk sect, String file, String url, StringBuffer buf, int no) throws Exception {
    	String saved_up_link = up_link;
    	boolean new_file = false;
    	if (buf == null) {
    		String file_path = out_dir + "/" + file + ".txt";
    		buf = new StringBuffer();
    		files.put(file_path, buf);
    		new_file = true;
    		buf.append("[["+contents_link+" | <contents>]] ");
    		if (prev_link != null && !(up_link != null && prev_link.equals(up_link)))
    			buf.append("[["+prev_link+" | <prev>]] ");
    		if (next_link_buf != null && next_link_pos >= 0)
    			next_link_buf.insert(next_link_pos, "[["+url+" | <next>]] ");
    		if (up_link != null)
    			buf.append("[["+up_link+" | <up>]] ");
    		up_link = prev_link = url;
    		next_link_buf = buf;
    		next_link_pos = buf.length();
    		buf.append("\n\n");
    	}
    	String sect_pad = "=";
    	for (int i = 0; i < sect.level; i++)
    		sect_pad += "=";
    	String title = null;
    	if (sect.title != null && sect.title.length() > 0)
    		title = sect.title;
    	else
    		title = "title" + no;
    	if (sect.level > 0) {
   			buf.append("\n"+sect_pad+" "+title+" "+sect_pad+"\n\n");
   			String numberer = "";
   			for (int j = 0; j < sect.level; j++)
   				numberer += "#";
   			contents_buf.append(numberer + " [["
   								+ url + (new_file ? "" : "#" + title)
   								+ " | " + title + "]]\n");
    	}
    	Chunk[] subs = sect.getSubs();
    	for (int i = 0; i < subs.length; i++) {
    		String file_step = "s" + (i+1);
    		String url_step = "/sect"+(i+1);
    		String sub_file = file + "_" + file_step;
    		if (SPLIT_SIZE != -1 && subs[i].totalLength() > SPLIT_SIZE) {
    			buf.append("\n[["+ url_step +" | "+subs[i].title+"]]\n");
    			outputSection(subs[i], sub_file, url + url_step, null, i+1);
    		} else {
    			outputSection(subs[i], sub_file, url, buf, i+1);
    		}
    	}
    	buf.append(sect.getText());
    	up_link = saved_up_link;
    }

	/*
     * ===================== Chunk ===========================
     */

	public static class Chunk {
		String title;
		int level;
		StringBuffer buf;
		Vector subs;
		public Chunk (String title, int level) {
			this.title = title;
			this.level = level;
			this.buf = new StringBuffer();
			this.subs = new Vector();
		}
		public void addSub(Chunk c) {
			subs.add(c);
		}
		public Chunk[] getSubs() {
			int n = subs.size();
			Chunk[] a = new Chunk[n];
			for (int i = 0; i < n; i++)
				a[i] = (Chunk) subs.get(i);
			return a;
		}
		public void append(String s) {
			buf.append(s);
		}
		public void append(char c) {
			buf.append(c);
		}
		public String getText() {
			return buf.toString();
		}
		public int selfLength() {
			return buf.length();
		}
		public int totalLength() {
			int len = selfLength();
			Chunk[] a = getSubs();
			for (int i = 0; i < a.length; i++)
				len += a[i].totalLength();
			return len;
		}
	}

    /*
     * ========================= Element ===============================
     */

	protected static final String inline_spans = "replaceable varname systemitem parameter";

    public void traverseElement(Element elem) throws Exception {
    	String name = elem.getNodeName();
    	if ("para".equals(name)) {
    		String pname = elem.getParentNode().getNodeName();
    		if ("entry".equals(pname)) {
    			Node prev = elem.getPreviousSibling();
    			if (prev != null && prev.getNodeName().equals("para"))
    				tag("<br>");
        		traverseChildren(elem);
        		return;
    		}
			Node prev = elem.getPreviousSibling();
			if (!"listitem".equals(pname))
				if (prev != null && prev.getNodeName().equals("para"))
					nl();
    		traverseChildren(elem);
    		nl();
    		return;
    	}
    	if (Util.isOneOf(name, inline_spans)) {
    		tag("<span class=\""+name+"\">");    		
    		traverseChildren(elem);
    		end("span");
    		return;
    	}
    	if ("emphasis".equals(name)) {
    		String role = elem.getAttribute("role");
    		String surround = "";
    		if ("italic".equals(role))
    			surround = "''";
    		else if ("".equals(role) || Util.isOneOf(role, "bold strong num"))
    			surround = "'''";
    		tag(surround);
    		traverseChildren(elem);
    		tag(surround);
    		return;
    	}
    	if ("listitem".equals(name)) {
    		handleListItem(elem);
    		return;
    	}
    	if ("subscript".equals(name)) {
    		start("sub");
    		traverseChildren(elem);
    		end("sub");
    		return;
    	}
    	if ("superscript".equals(name)) {
    		start("sup");
    		traverseChildren(elem);
    		end("sup");
    		return;
    	}
    	int level = -1;
    	if ("epigraph".equals(name))
    		level = 1;
    	else if ("sect1".equals(name))
    		level = 1;
    	else if ("sect2".equals(name))
    		level = 2;
    	else if ("sect3".equals(name))
    		level = 3;
    	else if ("sect4".equals(name))
    		level = 4;
    	else if ("sect5".equals(name))
    		level = 5;
    	else if ("sect6".equals(name))
    		level = 6;
    	if (level >= 0) {
    		String title;
    		if ("epigraph".equals(name))
    			title = "Epigraph";
    		else
    			title = getElementByTag(elem, "title").getTextContent();
    		makeSection(elem, title, level);
    		return;
    	}
    	if ("inlineequation".equals(name)) {
    		String alt = getElementByTag(elem, "alt").getTextContent();
    		start("math");
    		tag(convertStarMathToTex(alt));
    		end("math");
    		return;
    	}
    	if ("programlisting".equals(name)) {
    		boolean saved_pre_mode = pre_mode;
    		pre_mode = true;
    		add("\n");
    		traverseChildren(elem);
    		pre_mode = saved_pre_mode;
    		nl();
    		return;
    	}
    	if ("entry".equals(name)) {
    		tag(in_head ? "! " : "| ");
    		boolean saved_no_newlines = no_newlines;
    		no_newlines = true;
    		traverseChildren(elem);
    		no_newlines = saved_no_newlines;
    		tag("\n");
    	    return;
    	}
    	if ("tgroup".equals(name)) {
    		tag("{| border=\"1\" cellpadding=\"5\" cellspacing=\"0\" align=\"center\"\n");
    		traverseChildren(elem);
    	    tag("|}\n");
    	    return;
    	}
    	if ("thead".equals(name)) {
    		boolean save_in_head = in_head;
    		in_head = true;
    		traverseChildren(elem);
    	    in_head = save_in_head;
    	    return;
    	}
    	if ("tbody".equals(name)) {
    		Node prev = elem.getPreviousSibling();
    		if (prev != null && prev.getNodeName().equals("thead"))
    			tag("|-\n");
    		traverseChildren(elem);
    	    return;
    	}
    	if ("row".equals(name)) {
    		if (elem.getPreviousSibling() != null)
    			tag("|-\n");
    		traverseChildren(elem);
    	    return;
    	}
    	if ("inlinegraphic".equals(name) || "imagedata".equals(name)) {
    		String format = elem.getAttribute("format").trim();
    		if (Util.isImageFormat(format) && !"WMF".equalsIgnoreCase(format)) {
    			start("imagedata");
    			String out_ref = copyFile(elem.getAttribute("fileref"));
    			tag(out_ref);
    			end("imagedata");
    			String comment = elem.getAttribute("srccredit");
    			if (comment.startsWith("ERROR: ")) {
    				tag(" <i><small>");
    				add(comment);
    				tag("</small></i>\n");
    			}
    		}
    		return;
    	}
    	if ("math:math".equals(name) || "mml:math".equals(name)
    		|| "title".equals(name) || "articleinfo".equals(name))
    		return;
    	traverseChildren(elem);
    }

	protected void handleListItem (Element elem) throws Exception {
		StringBuffer prefix = new StringBuffer();
		Element top = content.getDocumentElement();
		Element parent = (Element) elem.getParentNode();
		boolean continued = false;
		while (parent != top) {
			String pn = parent.getNodeName();
			if ("itemizedlist".equals(pn)) {
				prefix.insert(0, '*');
			} else if ("orderedlist".equals(pn)) {
				prefix.insert(0, '#');
				if (parent.getAttribute("continuation").equals("continues"))
					continued = true;
			}
			parent = (Element)parent.getParentNode();
		}
		if (continued) {
			// this requires my own extension to the MediaWiki syntax...
			prefix.append('=');
		}
		prefix.append(' ');
		tag(prefix.toString());
		boolean saved_no_newlines = no_newlines;
		no_newlines = true;
		traverseChildren(elem);
		no_newlines = saved_no_newlines;
	}

	protected String convertStarMathToTex (String starmath) {
		String tex = smc.convert(starmath);
		// FIXME: below is a temporary workaround until bug 20 is fixed !
		if (tex.indexOf('\u2282') >= 0)
			tex = tex.replaceAll("\u2282", " \\\\subset ");
		return tex;
	}

	/*
     * ========================= Traverser ===============================
     */

	protected void traverseNode(Node node) throws Exception {
		switch(node.getNodeType()) {
		case Node.ELEMENT_NODE:
			Element elem = (Element)node;
			traverseElement(elem);
			break;
		case Node.TEXT_NODE:
			String text = node.getNodeValue();
			traverseText(text);
			break;
		}
	}

	protected void traverseChildren(Element elem) throws Exception {
		NodeList nl = elem.getChildNodes();
		for (int i = 0; i < nl.getLength(); i++)
			traverseNode(nl.item(i));
	}

	protected void traverseText(String text)  {
		add(text);
	}

	protected static final Element getElementByTag(Node in, String tag) {
		if (in != null && in.getNodeType() == Node.ELEMENT_NODE) {
			NodeList nl = ((Element)in).getElementsByTagName(tag);
			if (nl != null && nl.getLength() > 0)
				return (Element) nl.item(0);
		}
		return null;
	}

    /*
     * ========================= Utilities ===============================
     */
	
	protected String copyFile (String in_ref) throws IOException {
		String out_ref = in_ref;
		int pos = out_ref.lastIndexOf('/');
		if (pos >= 0)
			out_ref = out_ref.substring(pos+1);
		File in_file = new File(in_dir, in_ref);
		File out_file = new File(out_dir, out_ref);
		byte[] buf = new byte[8192];
		FileInputStream fis = new FileInputStream(in_file);
		FileOutputStream fos = new FileOutputStream(out_file);
		int n;
		while ((n = fis.read(buf)) >= 0)
			fos.write(buf, 0, n);
		fos.close();
		fis.close();
		return out_ref;
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
    
}

