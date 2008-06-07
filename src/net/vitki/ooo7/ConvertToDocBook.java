package net.vitki.ooo7;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.IOException;
import java.io.File;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Properties;

import net.vitki.wmf.Converter;

import org.w3c.dom.*;

/**
 * @author vit
 *
 */
public class ConvertToDocBook extends OooDocument
{
    public static boolean SENSE_ITALIC_SPANS = false;
	public static boolean SENSE_MATH_TEXT = false;
	public static boolean COPY_MATHML = false;
	public static boolean STRUCTURE_ONLY = false;
	public static boolean RENDER_WMF_METAFILES = false;
	public static boolean DROP_CONTENTS_SECTION = false;
	public static boolean SENSE_CIRCLED_NUMBERS = false;

    public static final String DOCBOOK_PUB_ID = "-//OASIS//DTD DocBook XML V4.3//EN";
    public static final String DOCBOOK_SYS_ID = "http://www.oasis-open.org/docbook/xml/4.3/docbookx.dtd";
	public static final String BROKENIMAGE_RESOURCE = "net/vitki/ooo7/brokenimage.png";
	public static final String BROKENIMAGE_FILENAME = "brokenimage.png";

    protected static final int SPAN_MODE = 0x10;

	private HashMap auto_styles;
	private HashMap list_styles;
	private Document out;
	private String data_dir;
	private int graphic_count;
	private boolean brokenimage_copied;

	public ConvertToDocBook() throws Exception {
		auto_styles = new HashMap();
		list_styles = new HashMap();
		out = null;
		data_dir = null;
		graphic_count = 0;
		brokenimage_copied = false;
	}

	protected static HashSet tags_to_skip;
	protected static String all_math_glyphs;

	private static void setupConstants() {
		tags_to_skip = new HashSet();
		tags_to_skip.add("text:variable-set");
		tags_to_skip.add("text:variable-get");
		tags_to_skip.add("text:bookmark-start");
		tags_to_skip.add("text:bookmark-end");
		tags_to_skip.add("text:footnote-citation");
		tags_to_skip.add("text:footnote");
		tags_to_skip.add("text:footnote-body");
		tags_to_skip.add("draw:text-box");
		tags_to_skip.add("text:a");
		tags_to_skip.add("text:line-break");
		tags_to_skip.add("text:tab-stop");
		tags_to_skip.add("text:reference-mark-start");
		tags_to_skip.add("text:reference-mark-end");
		tags_to_skip.add("comment");
		tags_to_skip.add("text:alphabetical-index-mark-start");
		tags_to_skip.add("text:alphabetical-index-mark-end");
		tags_to_skip.add("text:alphabetical-index");
		tags_to_skip.add("text:index-body");
		
		StringBuffer sb = new StringBuffer();
		int i;
		for (i = 0; i < 256; i++) {
			if (symbol_2_unicode[i] != 0) {
				sb.append((char)(0xf000 + i));
				sb.append((char)symbol_2_unicode[i]);
			}
		}
		for (i = 0x391; i <= 0x3d6; i++) // greek letters
			sb.append((char)i);
		all_math_glyphs = sb.toString();
	}

	/*
     * ========================= Styles ===============================
     */

	protected static class AutoStyleDesc {
		String  name;
		String  parent_name;
		boolean is_italic;
		boolean is_subscript;
		boolean is_superscript;
		boolean is_bold;
		String font_face;
		AutoStyleDesc(String name) {
			this.name = name;
			parent_name = null;
			is_subscript = is_superscript = false;
			is_bold = is_italic = false;
			font_face = null;
		}
	}

	protected void handleAutoStyle(String name, Element style) {
		AutoStyleDesc asd = new AutoStyleDesc(name);
		asd.parent_name = style.getAttribute("style:parent-style-name");
		if ("".equals(asd.parent_name))
			asd.parent_name = null;
		Element text_props = getElementByTag(style, "style:text-properties");
		if (text_props != null) {
			if (text_props.getAttribute("fo:font-weight").equals("bold"))
				asd.is_bold = true;
			if (text_props.getAttribute("fo:font-style").equals("italic"))
				asd.is_italic = true;
			else if (text_props.getAttribute("style:font-name").equals("Symbol")
						&& text_props.getAttribute("style:font-style-complex").equals("italic"))
				asd.is_italic = true;
			String position = text_props.getAttribute("style:text-position");
			if (position.startsWith("sub "))
				asd.is_subscript = true;
			else if (position.startsWith("super "))
				asd.is_superscript = true;
			asd.font_face = getAttributeOrNull(text_props, "style:font-name");
		}
		auto_styles.put(name, asd);
	}

	private final String getStyleName(String style_name) {
		String result = style_name;
		AutoStyleDesc asd = (AutoStyleDesc) auto_styles.get(style_name);
		if (asd != null && asd.parent_name != null)
			result = asd.parent_name;
		result = result.replace("_20_", " ");
		return result;
	}

	private final String getStyleName(Element elem) {
		String style_name = elem.getAttribute("text:style-name");
		return getStyleName(style_name);
	}

	private final AutoStyleDesc getAutoStyle(Element elem) {
		String style_name = elem.getAttribute("text:style-name");
		AutoStyleDesc asd;
		asd = (AutoStyleDesc) auto_styles.get(style_name);
		return asd;
	}

	protected boolean isTextOfTerm(String s) {
		// verify it contains only cyrillic letters or spaces
		int n = s.length();
		for (int i = 0; i < n; i++) {
			char c = s.charAt(i);
			if (Character.isWhitespace(c))
				continue;
			if (c >= 0x410 && c < 0x450)
				continue;
			return false;
		}
		return true;
	}

	protected boolean isTermSpan(Node node) {
		if (SENSE_ITALIC_SPANS && node != null && node.getNodeType() == Node.ELEMENT_NODE) {
			if ("text:span".equals(node.getNodeName())) {
				AutoStyleDesc asd = getAutoStyle((Element)node);
				if (asd != null && asd.is_italic) {
					if (!asd.is_subscript && !asd.is_superscript) {
						if (isTextOfTerm(node.getTextContent()))
							return true;
					}
				}
			}
		}
		return false;
	}

	protected boolean isInlineMath(Node node) {
		if (node == null)
			return false;
		int type = node.getNodeType();
		if (type == Node.ELEMENT_NODE && SENSE_ITALIC_SPANS) {
			if ("text:span".equals(node.getNodeName())) {
				AutoStyleDesc asd = getAutoStyle((Element)node);
				if (asd != null && asd.is_italic) {
					return !isTermSpan(node);
				}
			}
		}
		if (SENSE_MATH_TEXT) {
			String text = node.getTextContent();
			int n = text == null ? 0 : text.length();
			if (n == 0)
				return false;
			for (int i = 0; i < n; i++)
				if (all_math_glyphs.indexOf(text.charAt(i)) < 0)
					return false;
			return true;
		}
		return false;
	}

	protected boolean isSubscriptSpan(Node node) {
		if (node.getNodeType() == Node.ELEMENT_NODE && "text:span".equals(node.getNodeName())) {
			AutoStyleDesc asd = getAutoStyle((Element)node);
			if (asd != null && asd.is_subscript)
				return true;
		}
		return false;
	}

	protected boolean isSuperscriptSpan(Node node) {
		if (node.getNodeType() == Node.ELEMENT_NODE && "text:span".equals(node.getNodeName())) {
			AutoStyleDesc asd = getAutoStyle((Element)node);
			if (asd != null && asd.is_superscript)
				return true;
		}
		return false;
	}

	protected boolean isBold(Node node) {
		if (node.getNodeType() == Node.ELEMENT_NODE && "text:span".equals(node.getNodeName())) {
			AutoStyleDesc asd = getAutoStyle((Element)node);
			if (asd != null && asd.is_bold)
				return true;
		}
		return false;
	}

	protected boolean isItalic(Node node) {
		if (node.getNodeType() == Node.ELEMENT_NODE && "text:span".equals(node.getNodeName())) {
			AutoStyleDesc asd = getAutoStyle((Element)node);
			if (asd != null && asd.is_italic)
				return true;
		}
		return false;
	}

	/*
     * ========================= Article Info ===============================
     */

	protected static final String getMetaString(Document meta_doc, String tag) {
		Element elem = getElementByTag(meta_doc.getDocumentElement(), tag);
		if (elem != null) {
			String text = elem.getTextContent();
			if (!"".equals(text))
				return text;
		}
		return null;
	}

	private final Element appendDateElement (Element dst, String tag, String str) {
		if (str != null) {
			int pos = str.indexOf('T');
			if (pos > 0)
				str = str.substring(0, pos);
		}
		return appendTextElement(dst, tag, str);
	}

	protected void craftArticleInfo (Element article) throws Exception {
		if (STRUCTURE_ONLY)
			return;
		Element info = out.createElement("articleinfo");
		article.appendChild(info);
		Document meta = getDocumentBuilder().parse(getEntryStream("meta.xml"));
		appendTextElement(info, "title", getMetaString(meta, "dc:title"));
		String str = getMetaString(meta, "meta:initial-creator");
		if (str != null) {
			Element author = appendElement(info, "author");
			str = str.trim();
			int pos = str.indexOf(' ');
			if (pos < 0)
				appendTextElement(author, "othername", str);
			else {
				appendTextElement(author, "firstname", str.substring(0, pos).trim());
				appendTextElement(author, "surname", str.substring(pos+1).trim());
			}
		}
		appendDateElement(info, "pubdate", getMetaString(meta, "meta:creation-date"));
		appendDateElement(info, "date", getMetaString(meta, "dc:date"));
		appendTextElement(info, "issuenum", getMetaString(meta, "meta:editing-cycles"));
	}

	/*
     * ========================= Traverser ===============================
     */

	protected void traverseContent(Element src, int mode) throws Exception {
		Element article = out.createElement("article");
		out.appendChild(article);
		craftArticleInfo(article);
		traverseContent(src, article, mode);
	}

	protected void traverseContent(Element src, Element dst, int mode) throws Exception {
		String sname = src.getNodeName();
		if ("text:s".equals(sname)) {
			handleSpacer(src, dst, mode);
			return;
		}
		if ("draw:frame".equals(sname)) {
			handleFrame(src, dst, mode);
			return;
		}
		if (mode == SPAN_MODE)
			return;
		if ("text:h".equals(sname)) {
			handleHeading(src, dst, mode);
			return;
		}
		if ("text:p".equals(sname)) {
			handlePara(src, dst, mode);
			return;
		}
		if ("text:span".equals(sname)) {
			handleSpan(src, dst, mode);
			return;
		}
		if ("text:ordered-list".equals(sname) || "text:unordered-list".equals(sname)
				|| "text:list".equals(sname)) {
			handleList(src, dst, mode);
			return;
		}
		if ("text:list-item".equals(sname)) {
			handleListItem(src, dst, mode);
			return;
		}
		if ("table:table-cell".equals(sname)) {
			handleCell(src, dst, mode);
			return;	
		}
		if ("table:table-row".equals(sname)) {
			handleRow(src, dst, mode);
			return;
		}
		if ("table:table".equals(sname)) {
			handleTable(src, dst, mode);
			return;
		}
		if ("table:table-column".equals(sname)) {
			super.traverseContent(src, dst, mode);
			return;
		}
		if ("table:table-header-rows".equals(sname)) {
			Element thead = appendElement(dst, "thead");
			super.traverseContent(src, thead, mode);
			return;			
		}
		if (!tags_to_skip.contains(src))
			super.traverseContent(src, dst, mode);
	}

	protected void traverseContent(String text, Element dst, int mode) {
		if (mode != SPAN_MODE)
			dst.appendChild(out.createTextNode(symbolToUnicode(text)));
	}

	/*
     * ========================= Section ===============================
     */

	protected static final String subsection_styles = "Подраздел1 Подраздел2";

	protected int getOutlineLevel(Node node) {
		if (node == null || node.getNodeType() != Node.ELEMENT_NODE)
			return -1;
		String name = node.getNodeName();
		if (!"text:h".equals(name)) {
			if ("text:p".equals(name) &&
					Util.isOneOf(getStyleName((Element)node), subsection_styles))
				return 0;
			return -1;
		}
		try {
			return Integer.parseInt(((Element)node).getAttribute("text:outline-level"));
		} catch (Exception e) {
			return -1;
		}
	}
	
	int prev_level = 0;

	protected Node handleHeading(Element src, Element dst, int mode) throws Exception {
		int level = getOutlineLevel(src);
		int prev = prev_level;
		boolean incremental;
		if (level > 0) {
			prev_level = level;
			incremental = false;
		} else {
			level = prev_level + 1;
			if (level <= 0)
				level = 1;
			incremental = true;
		}
		if (prev <= 0) {
			prev = 0;
			NodeList nl = dst.getChildNodes();
			int n = nl.getLength();
			int i, j;
			for (i = 0; i < n; i++) {
				String iname = nl.item(i).getNodeName();
				if ("articleinfo".equals(iname))
					continue;
				break;
			}
			if (i < n) {
				Element epigraph = out.createElement("epigraph");
				for (j = i; j < n; j++)
					epigraph.appendChild(nl.item(j).cloneNode(true));
				for (j = n-1; j >= i; j--)
					dst.removeChild(nl.item(j));
				dst.appendChild(epigraph);
			}
		}
		for (int i = prev+1; i < level; i++)
			appendElement(dst = appendElement(dst, "sect"+i), "title");
		String title = src.getTextContent();
		boolean dropped = DROP_CONTENTS_SECTION && Util.isContentsTitle(title);
		Element sect = appendElement(dst, "sect"+level);
		if (dropped)
			dst.removeChild(sect);
		if (incremental)
			sect.setAttribute("status", "subsection");
		appendTextElement(sect, "title", title);
		Node node = src.getNextSibling();
		int count = 0;
		int next_level = -1;
		while(node != null && (next_level = getOutlineLevel(node)) < 0) {
			if (!dropped)
				traverseContent(node, sect, mode);
			beenHere(node);
			count++;
			node = node.getNextSibling();
		}
		if ((count == 0 || STRUCTURE_ONLY) && next_level <= level && !dropped)
			appendElement(sect, "para");
		do {
			Node saved_node = node;
			if (next_level == 0)
				next_level = incremental ? level : level + 1;
			if (next_level > level) {
				node = handleHeading((Element)node, sect, mode);
				beenHere(saved_node);
				next_level = getOutlineLevel(node);
			} else {
				if (next_level == level) {
					node = handleHeading((Element)node, dst, mode);
					beenHere(saved_node);
					next_level = getOutlineLevel(node);
				} else {
					break;
				}
			}
		} while(next_level >= level);
		return node;
	}

	/*
     * ========================= List ===============================
     */

	protected void handleList(Element src, Element dst, int mode) throws Exception {
		if (STRUCTURE_ONLY)
			return;
		String sname = src.getNodeName();
		String numeration = "arabic";
		char marker = 0; // none
		boolean ordered = true;
		if ("text:ordered-list".equals(sname))
			ordered = true;
		else if ("text:unordered-list".equals(sname))
			ordered = false;
		else {
			ListLevel ll = getListStyle(src);
			int type = ll == null ? LIST_TYPE_OTHER : ll.type;
			switch(type) {
			case LIST_TYPE_ARABIC:
				numeration = "arabic";
				break;
			case LIST_TYPE_LOWER_ROMAN:
				numeration = "lowerroman";
				break;
			case LIST_TYPE_UPPER_ROMAN:
				numeration = "upperroman";
				break;
			case LIST_TYPE_LOWER_ALPHA:
				numeration = "loweralpha";
				break;
			case LIST_TYPE_UPPER_ALPHA:
				numeration = "upperalpha";
				break;
			case LIST_TYPE_BULLETED:
				ordered = false;
				if (ll != null)
					marker = ll.marker;
				break;
			default:
				ordered = false;
				break;
			}
		}
		String tag = ordered ? "orderedlist" : "itemizedlist";
		boolean continued = false;
		if (ordered) {
			String cont_attr = src.getAttribute("text:continue-numbering");
			continued = "true".equals(cont_attr);
		}
		Element list = appendElement(dst, tag);
		if (continued)
			list.setAttribute("continuation", "continues");
		if (ordered && !"arabic".equals(numeration))
			list.setAttribute("numeration", numeration);
		if (!ordered && marker != 0 && marker != LIST_MARKER_BULLET)
			list.setAttribute("mark", ""+marker);
		super.traverseContent(src, list, mode);
	}

	protected void handleListItem(Element src, Element dst, int mode) throws Exception {
		Element item = appendElement(dst, "listitem");
		super.traverseContent(src, item, mode);
	}

	protected ListLevel getListStyle(Element list) {
		int level = 0;
		Element parent = list;
		Element top = getContentDocument().getDocumentElement();
		while (parent != top) {
			String pn = parent.getNodeName();
			if ("text:list".equals(pn)
					|| "text:ordered-list".equals(pn)
					|| "text:unordered-list".equals(pn))
				level++;
			parent = (Element)parent.getParentNode();
		}
		String style_name = list.getAttribute("text:style-name");
		return getListStyle(style_name, level);
	}

	protected ListLevel getListStyle(String style_name, int level) {
		return (ListLevel)list_styles.get(style_name + "--" +level);
	}

	protected void parseListStyles() throws Exception {
		Document style_doc = getStylesDocument();
		Element style_root = getElementByTag(style_doc.getDocumentElement(), "office:styles");
		NodeList ls_nl = style_root.getElementsByTagName("text:list-style");
		int ls_num = ls_nl.getLength();
		for (int i = 0; i < ls_num; i++) {
			Element ls = (Element) ls_nl.item(i);
			String ls_name = ls.getAttribute("style:name");
			NodeList ll_nl = ls.getElementsByTagName("*");
			int ll_num = ll_nl.getLength();
			for (int j = 0; j < ll_num; j++) {
				Element ll = (Element) ll_nl.item(j);
				int level;
				try {
					level = Integer.parseInt(ll.getAttribute("text:level"));
				} catch (Exception e) {
					level = 0;
				}
				String key = ls_name + "--" + level;
				int type = LIST_TYPE_OTHER;
				String ll_tag = ll.getLocalName();
				if ("list-level-style-number".equals(ll_tag)) {
					String num_format = ls.getAttribute("style:num-format");
					if ("1".equals(num_format))
						type = LIST_TYPE_ARABIC;
					else if ("i".equals(num_format))
						type = LIST_TYPE_LOWER_ROMAN;
					else if ("I".equals(num_format))
						type = LIST_TYPE_UPPER_ROMAN;
					else if ("a".equals(num_format))
						type = LIST_TYPE_LOWER_ALPHA;
					else if ("A".equals(num_format))
						type = LIST_TYPE_UPPER_ALPHA;
					else
						type = LIST_TYPE_ARABIC;
					list_styles.put(key, new ListLevel(type, (char)0));
				} else if ("list-level-style-bullet".equals(ll_tag)){
					char marker;
					try {
						marker = ls.getAttribute("text:bullet-char").charAt(0);
					} catch (Exception e) {
						marker = LIST_MARKER_BULLET;
					}
					list_styles.put(key, new ListLevel(LIST_TYPE_BULLETED, marker));
				}
			}
		}
	}

	public static final int LIST_TYPE_OTHER = 0;
	public static final int LIST_TYPE_BULLETED = -1;
	public static final int LIST_TYPE_ARABIC = 1;
	public static final int LIST_TYPE_LOWER_ROMAN = 2;
	public static final int LIST_TYPE_UPPER_ROMAN = 3;
	public static final int LIST_TYPE_LOWER_ALPHA = 4;
	public static final int LIST_TYPE_UPPER_ALPHA = 5;

	public static final char LIST_MARKER_BULLET = '\u2022';

	static class ListLevel {
		int type;
		char marker;
		ListLevel(int type, char marker) {
			this.type = type;
			this.marker = marker;
		}
	}

	/*
     * ========================= Paragraph ===============================
     */
	
	public static final String source_code_style = "Исходный код";

	protected void handlePara(Element src, Element dst, int mode) throws Exception {
		if (getOutlineLevel(src) >= 0) {
			handleHeading(src, dst, mode);
			return;
		}
		if (STRUCTURE_ONLY)
			return;
		if (getStyleName(src).equals(source_code_style)) {
			Element listing = appendElement(dst, "programlisting");
			while(true) {
				super.traverseContent(src, listing, mode);
				beenHere(src);
				Node next = src.getNextSibling();
				if (next.getNodeType() != Node.ELEMENT_NODE)
					break;
				src = (Element)next;
				if (!getStyleName(src).equals(source_code_style))
					break;
				listing.appendChild(out.createTextNode("\n"));
			}
			return;
		}
		Element para = appendElement(dst, "para");
		super.traverseContent(src, para, mode);
	}

	protected void handleSpacer(Element src, Element dst, int mode) throws Exception {
		int n;
		try {
			n = Integer.parseInt(src.getAttribute("text:c"));
		} catch (Exception e) {
			n = 0;
		}
		if (n > 0) {
			StringBuffer sb = new StringBuffer();
			for (int i = 0; i < n; i++)
				sb.append(' ');
			dst.appendChild(out.createTextNode(sb.toString()));		
		}
	}

	/*
     * ========================= Span ===============================
     */

	protected static HashMap inline_tag_map;

	protected static void setupInlineTagMapping() {
		inline_tag_map = new HashMap();
		inline_tag_map.put("Пункт определения", "systemitem");
		inline_tag_map.put("Мат.обозначение", "parameter");
		inline_tag_map.put("InlineFormulaStyle", "parameter");
		inline_tag_map.put("Шрифт имени", "citetitle");
	}

	protected void handleSpan(Element src, Element dst, int mode) throws Exception {
		String style = this.getStyleName(src);
		String tag = (String) inline_tag_map.get(style);

		AutoStyleDesc asd = getAutoStyle(src);
		if (SENSE_CIRCLED_NUMBERS
				&& asd != null && asd.font_face != null && !asd.is_italic
				&& !asd.is_subscript && !asd.is_superscript) {
			// smell out circled numbers
			boolean found = false;
			String num = null;
			if (asd.is_bold && "Arial Narrow".equals(asd.font_face)) {
				num = symbolToUnicode(src.getTextContent());
				int len = num.length();
				found = (len >= 3 && num.charAt(0) == '(' && num.charAt(len-1) == ')');
			} else if ("Wingdings".equals(asd.font_face)) {
				String text = symbolToUnicode(src.getTextContent());
				StringBuffer buf = new StringBuffer(text.length());
				for (int i = 0; i < text.length(); i++) {
					char c = text.charAt(i);
					if (c >= 0xf081 && c <= 0xf08a) {
						buf.append("(" + String.valueOf(c - 0xf080) + ")");
						found = true;
					} else {
						buf.append(c);
					}
				}
				num = buf.toString();
			}
			if (num != null && found) {
				appendTextElement(dst, "emphasis", num).setAttribute("role", "num");
				return;
			}
		}

		if (tag == null && isInlineMath(src)) {
			Element replaceable = appendElement(dst, "replaceable");
			Node node = src;
			Element subscript = null, superscript = null;
			do {
				String text = symbolToUnicode(node.getTextContent());
				if (isSubscriptSpan(node)) {
					if (subscript == null)
						subscript = appendTextElement(replaceable, "subscript", text);
					else
						subscript.appendChild(out.createTextNode(text));
					superscript = null;
				} else if (isSuperscriptSpan(node)) {
					if (superscript == null)
						superscript = appendTextElement(replaceable, "superscript", text);
					else
						superscript.appendChild(out.createTextNode(text));
					subscript = null;
				} else {
					replaceable.appendChild(out.createTextNode(text));
					subscript = superscript = null;
				}
				if (node.getNodeType() == Node.ELEMENT_NODE)
					super.traverseContent((Element)node, dst, SPAN_MODE);
				beenHere(node);
				node = node.getNextSibling();
			} while (isInlineMath(node));
			// replaceables is a wide notion. sometimes these are simple parameters.
			// let's find out.
			String text = replaceable.getTextContent().trim();
			boolean has_letters = false;
			for (int i = 0; i < text.length(); i++) {
				char c = text.charAt(i);
				if ((c>='a' && c<='z') || (c>='A' && c<='Z')) {
					has_letters = true;
					continue;
				}
				if ((c>='0' && c<='9') || c=='.' || c=='_')					
					continue;
				return;
			}
			if (!has_letters)
				return;
			// this is parameter, actually
			Element parameter = appendElement(dst, "parameter");
			NodeList nl = replaceable.getChildNodes();
			for (int i = 0; i < nl.getLength(); i++)
				parameter.appendChild(nl.item(i).cloneNode(true));
			replaceable.getParentNode().removeChild(replaceable);
			return;
		}
		Element span = dst;

		if (tag == null && isTermSpan(src)) {
			span = appendTextElement(dst, "firstterm", src.getTextContent());
			return;
		}

		if (tag == null && asd != null && asd.is_bold) {
			span = appendElement(dst, (tag = "emphasis"));
			span.setAttribute("role", "bold");
		} else if (tag == null && asd != null && asd.is_italic) {
			span = appendElement(dst, (tag = "emphasis"));
			span.setAttribute("role", "italic");
		} else if (tag != null) {
			span = appendElement(dst, tag);
			if ("citetitle".equals(tag))
				span.setAttribute("pubwork", "refentry");
		}

		Element subscript = null, superscript = null;
		do {
			String text = symbolToUnicode(src.getTextContent());
			if (isSubscriptSpan(src)) {
				if (subscript == null)
					subscript = appendTextElement(span, "subscript", text);
				else
					subscript.appendChild(out.createTextNode(text));
				superscript = null;
			} else if (isSuperscriptSpan(src)) {
				if (superscript == null)
					superscript = appendTextElement(span, "superscript", text);
				else
					superscript.appendChild(out.createTextNode(text));
				subscript = null;
			} else {
				span.appendChild(out.createTextNode(text));
				subscript = superscript = null;
			}
			super.traverseContent(src, dst, SPAN_MODE);
			beenHere(src);
			Node node = src.getNextSibling();
			if (node == null || node.getNodeType() != Node.ELEMENT_NODE)
				break;
			src = (Element)node;
		} while ("text:span".equals(src.getNodeName()) && getStyleName(src).equals(style));
	}

	/*
     * ========================= Table ===============================
     */

	protected void handleTable(Element src, Element dst, int mode) throws Exception {
		if (STRUCTURE_ONLY)
			return;
		if ("entry".equals(dst.getNodeName()))
			dst = appendElement(dst, "para");
		Element table = appendElement(dst, "informaltable");
		table.setAttribute("frame", "all");
		int cells = countDescendantElementsByTag(src, "table:table-cell");
		int rows = countDescendantElementsByTag(src, "table:table-row");
		int cols = cells / rows;
		int numcols = cols;
		NodeList cols_nl = src.getElementsByTagName("table:table-column");
		for (int i = 0; i < cols_nl.getLength(); i++) {
			String ncr = getAttributeOrNull(cols_nl.item(i), "table:number-columns-repeated");
			if (ncr != null) {
				numcols = Integer.parseInt(ncr);
				break;
			}
		}
		Element tgroup = appendElement(table, "tgroup");
		tgroup.setAttribute("cols", String.valueOf(numcols));
		for (int i = 0; i < numcols; i++) {
			int no = i + 1;
			Element colspec = appendElement(tgroup, "colspec");
			colspec.setAttribute("colnum", ""+no);
			colspec.setAttribute("colname", "c"+no);
		}
		super.traverseContent(src, tgroup, mode);
	}

	protected void handleRow(Element src, Element dst, int mode) throws Exception {
		String pname = src.getParentNode().getNodeName();
		if ("table:table-header-rows".equals(pname)) {
			Element row = appendElement(dst, "row");
			super.traverseContent(src, row, mode);
			return;
		}
		if (!"table:table".equals(pname)) {
			super.traverseContent(src, dst, mode);
			return;
		}
		Element tbody = getElementByTag(dst, "tbody");
		if (tbody == null)
			tbody = appendElement(dst, "tbody");
		Element row = appendElement(tbody, "row");
		super.traverseContent(src, row, mode);
	}

	protected static final int getSpannedColumns(Node node) {
		if (node == null || !node.getNodeName().equals("table:table-cell"))
			return 0;
		String attr = ((Element)node).getAttribute("table:number-columns-spanned").trim();
		if ("".equals(attr))
			return 1;
		return Integer.parseInt(attr);
	}

	protected void handleCell(Element src, Element dst, int mode) throws Exception {
		Element entry = appendElement(dst, "entry");
		int ncs = getSpannedColumns(src);
		if (ncs > 1) {
			int ncs_before = 0;
			Node node = src.getPreviousSibling();
			while(node != null) {
				ncs_before += getSpannedColumns(node);
				node = node.getPreviousSibling();
			}
			entry.setAttribute("namest", "c"+ncs_before);
			entry.setAttribute("nameend", "c"+(ncs_before+ncs));
		}
		super.traverseContent(src, entry, mode);
	}

	protected static final int countDescendantElementsByTag(Element elem, String tag) {
		return countDescendantElementsByTag(elem, tag, 0);		
	}

	private static final int countDescendantElementsByTag(Element elem, String tag, int sum) {
		if (elem.getNodeName().equals(tag))
			sum++;
		NodeList nl = elem.getChildNodes();
		int n = nl.getLength();
		for (int i = 0; i < n; i++) {
			if (nl.item(i).getNodeType() == Node.ELEMENT_NODE)
				sum = countDescendantElementsByTag((Element)nl.item(i), tag, sum);
		}
		return sum;
	}

	/*
     * ========================= Drawing ===============================
     */

	protected void handleFrame(Element src, Element dst, int mode) throws Exception {
		graphic_count++;
		Element img_el = getElementByTag(src, "draw:image");
		Element obj_el = getElementByTag(src, "draw:object");
		Element ole_el = getElementByTag(src, "draw:object-ole");
		String img_ref = Util.normalRef(getAttributeOrNull(img_el, "xlink:href"));
		String obj_ref = Util.normalRef(getAttributeOrNull(obj_el, "xlink:href"));
		String ole_ref = Util.normalRef(getAttributeOrNull(ole_el, "xlink:href"));
		int width = sizeToPixels(src.getAttribute("svg:width"));
		int height = sizeToPixels(src.getAttribute("svg:height"));
		if (obj_ref != null
				&& getMimeType(obj_ref+"/").equals("application/vnd.oasis.opendocument.formula")) {
			Element eqn_el = appendElement(dst, "inlineequation");
			InputStream mml_is = getEntryStream(obj_ref + "/content.xml");
			Document mml_doc = getDocumentBuilder().parse(withoutDtd(mml_is));
			mml_is.close();
			Element math_el = mml_doc.getDocumentElement();
			Element sema_el = getElementByTag(math_el, "math:semantics");
			Element anno_el = getElementByTag(sema_el, "math:annotation");
			Element alt_el = appendTextElement(eqn_el, "alt", anno_el.getTextContent());
			String anno_enc = getAttributeOrNull(anno_el, "math:encoding");
			if (anno_enc != null && !"StarMath 5.0".equals(anno_enc))
				alt_el.setAttribute("vendor", anno_enc);
			if (COPY_MATHML)
				eqn_el.appendChild(out.importNode(math_el, true));
			if (img_ref != null)
				appendGraphic(eqn_el, "graphic", "eqn", graphic_count, img_ref, width, height);
		} else if (ole_ref != null) {
			Element media = appendElement(dst, "inlinemediaobject");
			appendGraphic(media, "imagedata", "ole", graphic_count, ole_ref, width, height);
			if (img_ref != null)
				appendGraphic(media, "imagedata", "pic", graphic_count, img_ref, width, height);
		} else {
			if (RENDER_WMF_METAFILES && "WMF".equals(imageFormat(img_ref))) {
				Element media = appendElement(dst, "inlinemediaobject");
				appendGraphic(media, "imagedata", "pic", graphic_count, img_ref, width, height);
			} else {
				appendGraphic(dst, "inlinegraphic", "pic", graphic_count, img_ref, width, height);
			}
		}
	}

	protected void copyStreamToFile (InputStream is, String path) throws IOException {
		OutputStream os = new FileOutputStream(path);
		byte[] buf = new byte[32768];
		int n;
		while((n = is.read(buf)) > 0)
			os.write(buf, 0, n);
		is.close();
		os.close();
	}

	private Element appendGraphic(Element dst, String tag, String basename, int no,
									String ref,	int width, int height) throws Exception {
		String fmt = imageFormat(ref);
		String width_str = String.valueOf(width) + "px";
		String ext;
		if (Util.isImageFormat(fmt))
			ext = fmt.toLowerCase();
		else if ("PowerPoint".equals(fmt))
			ext = "ppt";
		else if (fmt.startsWith("Visio"))
			ext = "vis";
		else
			ext = "dat";
		String file_name = Util.num000(no) + basename + "." + ext;
		String path = data_dir + "/" + file_name;
		int pos = data_dir.lastIndexOf('/');
		String rel_dir = "./" + (pos < 0 ? data_dir : data_dir.substring(pos + 1)) + "/";
		copyStreamToFile(getEntryStream(ref), path);
		Element to = dst;
		if ("imagedata".equals(tag))
			to = appendElement(dst, "imageobject");
		Element elem = appendElement(to, tag);
		elem.setAttribute("fileref", rel_dir + file_name);
		if (Util.isImageFormat(fmt))
			elem.setAttribute("format", fmt);
		else {
			elem.setAttribute("format", "linespecific");
			elem.setAttribute("srccredit", fmt);
		}
		elem.setAttribute("width", width_str);
		if ("WMF".equals(fmt) && RENDER_WMF_METAFILES) {
			String pic_name = Util.num000(no) + "wmf.png";
			String pic_path = data_dir + "/" + pic_name;
			String ermes = null;
	        try {
		        Properties props = Converter.getDefaultProperties();
		        props.setProperty ("width", String.valueOf(width));
		        props.setProperty ("height", String.valueOf(height));
		        props.setProperty ("ignore-quirks", "yes");
	        	InputStream wmf_is = new FileInputStream(path);
	        	OutputStream png_os = new FileOutputStream(pic_path);
	        	Converter.convert (wmf_is, png_os, props);
	        	wmf_is.close();
	        	png_os.close();
	        } catch (Exception e) {
	        	ermes = e.toString();
	        	(new File(pic_path)).delete();
	        	System.err.println("cannot render \""+path+"\": "+ermes);
	        	pic_name = BROKENIMAGE_FILENAME;
	    		pic_path = data_dir + "/" + pic_name;
	    		if (!brokenimage_copied) {
		    		ClassLoader cl = this.getClass().getClassLoader();
		    		copyStreamToFile(cl.getResourceAsStream(BROKENIMAGE_RESOURCE), pic_path);
	    		}
			}
	        if ("imagedata".equals(tag))
	        	to = appendElement(dst, "imageobject");
	        Element pic = appendElement(to, tag);
	        pic.setAttribute("fileref", rel_dir + pic_name);
	        pic.setAttribute("format", "PNG");
	        pic.setAttribute("width", width_str);
	        pic.setAttribute("srccredit", ermes != null ? "ERROR: "+ermes : "automade");
		}
		return elem;
	}

	public static final double INCHES_2_PIXELS = 145.68;
	public static final double CENTIMETERS_2_PIXELS = INCHES_2_PIXELS / 2.54;
	public static final double MILLIMETERS_2_PIXELS = CENTIMETERS_2_PIXELS / 10.0;
	public static final double POINTS_2_PIXELS = INCHES_2_PIXELS / 72.0;
	public static final double PICAS_2_PIXELS = INCHES_2_PIXELS / 6.0;

	protected static final int sizeToPixels (String str) {
		str = str.trim();
		String units = str.substring(str.length()-2);
		double val;
		try {
			val = Double.parseDouble(str.substring(0, str.length()-2).trim());
		} catch (Exception e) {
			val = 0;
		}
		double coef = 0;
		if ("in".equals(units))
			coef = INCHES_2_PIXELS;
		else if ("px".equals(units))
			coef = 1;
		else if ("cm".equals(units))
			coef = CENTIMETERS_2_PIXELS;
		else if ("mm".equals(units))
			coef = MILLIMETERS_2_PIXELS;
		else if ("pt".equals(units))
			coef = POINTS_2_PIXELS;
		else if ("pc".equals(units))
			coef = PICAS_2_PIXELS;
		if (val == 0 || coef == 0)
			throw new RuntimeException("Incorrect size "+str);
		return (int)(val * coef + 0.5);
	}

	protected String imageFormat(String ref) {
		if (ref == null || "".equals(ref))
			return "unknown";
		ref = Util.normalRef(ref);
		int pos = ref.lastIndexOf('.');
		String ext = pos < 0 || pos == ref.length()-1 ? "" : ref.substring(pos+1);
		ext = ext.toUpperCase();
		String fmt = ext;
		if (!Util.isImageFormat(ext)) {
			if (ref.startsWith("ObjectReplacements/"))
				return "WMF";
			fmt = getMimeType(ref);
			if (fmt == null || "".equals(fmt)
					|| "application/vnd.sun.star.oleobject".equals(fmt)) {
				try {
					fmt = getObjectType(ref).trim();
					if (fmt.toLowerCase().startsWith("ok"))
						fmt = fmt.substring(2).trim();
					if (fmt.toLowerCase().startsWith("microsoft ") && fmt.length() > 10)
						fmt = fmt.substring(10).trim();
				} catch (Exception e) {
					fmt = null;
				}
				if (fmt == null || "".equals(fmt)
						|| "<TOO_SMALL_OLE>".equals(fmt) || "<MAGIC_NOT_FOUND>".equals(fmt))
					fmt = "unknown";
			}
		}
		return fmt;
	}

    /*
     * ============================= Utilities ============================
     */

	protected final Element appendTextElement (Element dst, String tag, String str) {
		Element elem = out.createElement(tag);
		if (str != null)
			elem.setTextContent(str);
		dst.appendChild(elem);
		return elem;
	}

	protected final Element appendElement (Element dst, String tag) {
		Element elem = out.createElement(tag);
		dst.appendChild(elem);
		return elem;
	}

    protected final String getAttributeOrNull (Node node, String name) {
    	if (node == null || node.getNodeType() != Node.ELEMENT_NODE)
    		return null;
    	String attr = ((Element)node).getAttribute(name);
    	return "".equals(attr) ? null : attr;
    }

	protected static final String symbolToUnicode(String s) {
		int n = s.length();
		StringBuffer b = new StringBuffer(n);
		boolean changed = false;
		for (int i = 0; i < n; i++) {
			char c = s.charAt(i);
			if (c >= '\uf000' && c <='\uf0ff') {
				c = (char)symbol_2_unicode[c - '\uf000'];
				if (c == 0)
					c = s.charAt(i);
				changed = true;
			}
			b.append(c);
		}
		return changed ? b.toString() : s;
	}

    private static int[] symbol_2_unicode = {
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x00
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x08
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x10
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x18
        0,      0x0021, 0x2200, 0x0023, 0x2203, 0x0025, 0x0026, 0x220d, // 0x20
        0x0028, 0x0029, 0x002a, 0x002b, 0x002c, 0x002d, 0x002e, 0x002f, // 0x28
        0x0030, 0x0031, 0x0032, 0x0033, 0x0034, 0x0035, 0x0036, 0x0037, // 0x30
        0x0038, 0x0039, 0x003a, 0x003b, 0x003c, 0x003d, 0x003e, 0x003f, // 0x38
        0x2245, 0x0391, 0x0392, 0x03a7, 0x0394, 0x0395, 0x03a6, 0x0393, // 0x40
        0x0397, 0x0399, 0x03d1, 0x039a, 0x039b, 0x039c, 0x039d, 0x039f, // 0x48
        0x03a0, 0x0398, 0x03a1, 0x03a3, 0x03a4, 0x03a5, 0x03c2, 0x03a9, // 0x50
        0x039e, 0x03a8, 0x0396, 0x005b, 0x2234, 0x005d, 0x22a5, 0x005f, // 0x58
        0x00af, 0x03b1, 0x03b2, 0x03c7, 0x03b4, 0x03b5, 0x03d5, 0x03b3, // 0x60
        0x03b7, 0x03b9, 0x03c6, 0x03ba, 0x03bb, 0x03bc, 0x03bd, 0x03bf, // 0x68
        0x03c0, 0x03b8, 0x03c1, 0x03c3, 0x03c4, 0x03c5, 0x03d6, 0x03c9, // 0x70
        0x03be, 0x03c8, 0x03b6, 0x007b, 0x007c, 0x007d, 0x007e, 0,      // 0x78
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x80
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x88
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x90
        0,      0,      0,      0,      0,      0,      0,      0,      // 0x98
        0,      0x03d2, 0x2032, 0x2264, 0x2044, 0x221e, 0x0192, 0x2663, // 0xa0
        0x2666, 0x2665, 0x2660, 0x2194, 0x2190, 0x2191, 0x2192, 0x2193, // 0xa8
        0x00b0, 0x00b1, 0x2033, 0x2265, 0x00d7, 0x221d, 0x2202, 0x2022, // 0xb0
        0x00f7, 0x2260, 0x2261, 0x2248, 0x2026, 0x2502, 0x2500, 0x21b5, // 0xb8
        0x2135, 0x2111, 0x211c, 0x2118, 0x2297, 0x2295, 0x2298, 0x2229, // 0xc0
        0x222a, 0x2283, 0x2287, 0x2284, 0x2282, 0x2286, 0x2208, 0x2209, // 0xc8
        0x2220, 0x2207, 0x00ae, 0x00a9, 0x2122, 0x220f, 0x221a, 0x2219, // 0xd0
        0x00ac, 0x2227, 0x2228, 0x21d4, 0x21d0, 0x21d1, 0x21d2, 0x21d3, // 0xd8
        0x25ca, 0x2329, 0x00ae, 0x00a9, 0x2122, 0x2211, 0x239b, 0x239c, // 0xe0
        0x239d, 0x23a1, 0x23a2, 0x23a3, 0x23a7, 0x23a8, 0x23a9, 0x23aa, // 0xe8
        0,      0x232a, 0x222b, 0x2320, 0x23ae, 0x2321, 0x239e, 0x239f, // 0xf0
        0x23a0, 0x23a4, 0x23a5, 0x23a6, 0x23ab, 0x23ac, 0x23ad, 0       // 0xf8
       };


    /*
     * ============================= main ============================
     */
	
    public static void main(String[] args) throws Exception {
    	String in_name = null;
    	String out_name = null;
    	String dbg_path = null;
    	boolean bad_usage = false;
    	for (int i = 0; i < args.length; i++) {
    		String opt = args[i];
    		if ("-out".equals(opt) || "-o".equals(opt)) {
    			out_name = args[++i];
    		} else if ("-debug".equals(opt) || "-d".equals(opt)) {
    			dbg_path = args[++i];
    		} else if ("-nodebug".equals(opt) || "-d-".equals(opt)) {
    			dbg_path = null;
    		} else if ("-indent".equals(opt) || "-I".equals(opt)) {
    			INDENTATION = true;
    		} else if ("-noindent".equals(opt) || "-I-".equals(opt)) {
    			INDENTATION = false;
    		} else if ("-iso".equals(opt)) {
    			ENCODING = "ISO-8859-1";
    		} else if ("-utf".equals(opt) || "-u8".equals(opt)) {
    			ENCODING = "UTF8";
    		} else if ("-progress".equals(opt) || "-h".equals(opt)) {
    			SHOW_PROGRESS = true;
    		} else if ("-noprogress".equals(opt) || "-h-".equals(opt)) {
    			SHOW_PROGRESS = false;
    		} else if ("-sense".equals(opt) || "-S".equals(opt)) {
    		    SENSE_ITALIC_SPANS = true;
    			SENSE_MATH_TEXT = true;
    		} else if ("-nosense".equals(opt) || "-S-".equals(opt)) {
    		    SENSE_ITALIC_SPANS = false;
    			SENSE_MATH_TEXT = false;
    		} else if ("-circlednumbers".equals(opt) || "-CN".equals(opt)) {
    		    SENSE_CIRCLED_NUMBERS = true;
    		} else if ("-nocirclednumbers".equals(opt) || "-CN-".equals(opt)) {
    		    SENSE_CIRCLED_NUMBERS = false;
    		} else if ("-mml".equals(opt) || "-M".equals(opt)) {
    			COPY_MATHML = true;
    		} else if ("-nomml".equals(opt) || "-M-".equals(opt)) {
    			COPY_MATHML = false;
    		} else if ("-dropcontents".equals(opt)) {
    			DROP_CONTENTS_SECTION = true;
    		} else if ("-nodropcontents".equals(opt)) {
    			DROP_CONTENTS_SECTION = false;
    		} else if ("-structure".equals(opt)) {
    			STRUCTURE_ONLY = true;
    		} else if ("-full".equals(opt)) {
    			STRUCTURE_ONLY = false;
    		} else if ("-render".equals(opt) || "-W".equals(opt)) {
    			RENDER_WMF_METAFILES = true;
    		} else if ("-norender".equals(opt) || "-W-".equals(opt)) {
    			RENDER_WMF_METAFILES = false;
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
			System.err.println("usage: java -jar ConvertToDocBook.jar"
								+" [-mml] [-sense] [-render] [-indent] [-structure]"
								+" [-dropcontents] [-circlednumbers]"
								+" [-debug debug_dir|+] [-iso | -utf] [-progress]"
								+" [-out <out.xml>] <in.odt> ");
			System.exit(1);
		}
		progressPrintln("setting up...");
		ConvertToDocBook ooo2db = new ConvertToDocBook();
		ooo2db.run(in_name, out_name, dbg_path);
	}
    
	public void run (String in_path, String out_path, String dbg_path) throws Exception {
		setInputPath(in_path);
		String basename = Util.getFileName(in_path); 
		if (out_path == null)
			out_path = basename +"_db.xml";
		setupLogging(dbg_path);
		data_dir = basename + "_files";
		Util.cleanDirectory(data_dir);
		System.err.println("start: " + in_path);
		progressPrint("parsing... ");
		parseInputDocument();
		parseListStyles();
		out = newDocument();
		traverseContent();
		progressPrintln(" OK");
		FileOutputStream fos = new FileOutputStream(out_path);
		writeDoc(fos, out, DOCBOOK_PUB_ID, DOCBOOK_SYS_ID);
		fos.close();
		closeLog();
		System.err.println("done: " + out_path);
	}

	static {
		setupConstants();
		setupInlineTagMapping();
	}
	
}

