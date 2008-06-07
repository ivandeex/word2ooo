package net.vitki.ooo7;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Iterator;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.*;

/**
 * @author vit
 *
 */
public class OooFormula
{
	public static final String MATH_NS = "http://www.w3.org/1998/Math/MathML";

	public static final String SIZING_ATTR = "scriptlevel";
	public static final int SIZING_BASE = 6;
	public static final int SIZING_FACTOR = 2; // StarMath_size = scriptlevel * factor + base

	public static boolean ADVANCED_ID_MERGING = true;
	public static boolean ADVANCED_TRIVIAL_SPACING = true;
	public static boolean ADVANCED_TRIVIAL_SUBSCRIPTING = true;

	private Document mml_doc;
	private StringBuffer buf;
	private String annotation;
	private boolean left_aligned;
	private boolean single_column;
	private char anno_last;

	public OooFormula() {
		buf = new StringBuffer();
		anno_last = 0;
		anno_last = 0;
		left_aligned = false;
		single_column = false;
		mml_doc = null;
	}
	
	public OooFormula(Element orig_mml) throws Exception {
		this();
		refactorMath(orig_mml);
	}
	
	public String getFake() {
		return (String)fakes.get(annotation);
	}
	
	public String getAnnotation() {
		return annotation;
	}
	
	public Document getDocument() {
		return mml_doc;
	}
	
	public Document refactorMath (Element orig_mml) throws Exception {
		mml_doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
		Element root = mml_doc.createElementNS(MATH_NS, "math:math");
		mml_doc.appendChild(root);
		transformChildren(orig_mml, root);
		annotateMath(root);
		Element anno_el = mml_doc.createElementNS(MATH_NS, "math:annotation");
		anno_el.setAttributeNS(MATH_NS, "math:encoding", "StarMath 5.0");
		annotation = buf.toString();
		buf.setLength(0);
		anno_el.setTextContent(annotation);
		root.getElementsByTagName("math:semantics").item(0).appendChild(anno_el);
		return mml_doc;
	}

    /*
     * ========================= Translations ===============================
     */

	private static boolean translations_ok = false;
	
	private static HashMap backops;
	private static HashSet binary_ops;
	private static HashMap ops;
	private static HashMap syms;
	private static HashMap summings;
	private static HashMap brackets;
	private static HashMap scalable;
	private static HashMap modifiers;
	private static HashMap fakes;

	private static final String brackets_glyphs = "()[]{}";
	private static final String prime_glyphs = "'\u2032";
	private static final String delim_glyphs = "...,;";
	private static final String MATH_SYM_FUNC = "\u2061";
	private static final String MATH_ASSIGNMENT = ":="; // Is there a better glyph ?
	private static final String PHANTOM_OPERAND = "`";
	
	static {
		setupMathTranslations();
	}

	private static void setupMathTranslations() {
		if (translations_ok)
			return;
		backops = new HashMap();
		ops = new HashMap();
		binary_ops = new HashSet();
		syms = new HashMap();
		summings = new HashMap();
		brackets = new HashMap();
		scalable = new HashMap();
		modifiers = new HashMap();
		fakes = new HashMap();

		ops.put("\ue083", "+");
		ops.put("\u2212", "-");
		ops.put("\u00b1", "+-");
		ops.put("\u2260", "<>");
		ops.put("\u2264", "<=");
		ops.put("\u2265", ">=");
		ops.put("\u2192", "%tendto");
		ops.put("\u2208", "in");
		ops.put("\u2209", "notin");
		ops.put("\u2200", "forall");
		ops.put("\u2203", "exists");
		ops.put("\u2204", "\u2204"); // ??? the "DoesNotExist" operator
		ops.put("\u22c5", "cdot");
		ops.put("\u2213", "-+");
		ops.put("\u2227", "and");
		ops.put("\u2228", "or");
		ops.put("\u2229", "intersection");
		ops.put("\u222a", "union");
		ops.put("\u2202", "partial"); // partial differential (is it a symbol ?)
		ops.put("\u226a", "<<");
		ops.put("\u226b", ">>");
		ops.put(MATH_SYM_FUNC, ""); // function application
		ops.put("|", "divides");
		ops.put("\u22ef", "dotsaxis"); // midline horizontal ellipsis
		ops.put("\u2205", "emptyset");
		ops.put("\u2261", "equiv");
		ops.put("\u220f", "{prod{}}");	// production
		ops.put("\u2211", "{sum{}}");	// sum
		ops.put("\u222b", "{int{}}");	// integral

		summings.put("\u220f", "prod");
		summings.put("{prod{}}", "prod");
		summings.put("\u2211", "sum");
		summings.put("{sum{}}", "sum");
		summings.put("\u222b", "int");
		summings.put("{int{}}", "int");

		binary_ops.add("+");
		//binary_ops.add("-"); // can be unary
		binary_ops.add("*");
		binary_ops.add("/");
		binary_ops.add("wideslash");
		binary_ops.add("+-");
		binary_ops.add("-+");
		binary_ops.add("<>");
		binary_ops.add("<");
		binary_ops.add(">");
		binary_ops.add("=");
		binary_ops.add("<=");
		binary_ops.add(">=");
		binary_ops.add("in");
		binary_ops.add("notin");
		binary_ops.add("cdot");
		binary_ops.add("and");
		binary_ops.add("or");
		binary_ops.add("intersection");
		binary_ops.add("union");
		binary_ops.add("%strictlylessthan");
		binary_ops.add("%strictlygreaterthan");
		binary_ops.add("divides");
		binary_ops.add("equiv");

		reverseHashMap(ops, backops);

		brackets.put("\u2308", "left lceil");
		brackets.put("\u2309", "right rceil");
		brackets.put("\u230a", "left lfloor");
		brackets.put("\u230b", "right rfloor");
		brackets.put("\u2329", "left langle");
		brackets.put("\u232a", "right rangle");
		brackets.put("(", "(");
		brackets.put(")", ")");
		brackets.put("[", "[");
		brackets.put("]", "]");
		brackets.put("{", "left lbrace");
		brackets.put("}", "right rbrace");

		scalable.put("(", "left(");
		scalable.put(")", "right)");
		scalable.put("[", "left[");
		scalable.put("]", "right]");
		scalable.put("{", "left lbrace");
		scalable.put("}", "right rbrace");

		modifiers.put("\u223c", "widetilde");
		modifiers.put("\u02dc", "tilde");
		modifiers.put("\u00af", "overline");
		modifiers.put("^", "hat");

		// capital greek letters 
		syms.put("\u0391", "%ALPHA");
		syms.put("\u0392", "%BETA");
		syms.put("\u0393", "%GAMMA");
		syms.put("\u0394", "%DELTA");
		syms.put("\u0395", "%EPSILON");
		syms.put("\u0396", "%ZETA");
		syms.put("\u0397", "%ETA");
		syms.put("\u0398", "%THETA");
		syms.put("\u0399", "%IOTA");
		syms.put("\u039a", "%KAPPA");
		syms.put("\u039b", "%LAMBDA");
		syms.put("\u039c", "%MU");
		syms.put("\u039d", "%NU");
		syms.put("\u039e", "%XI");
		syms.put("\u039f", "%OMICRON");
		syms.put("\u03a0", "%PI");
		syms.put("\u03a1", "%RHO");
		syms.put("\u03a3", "%SIGMA");
		syms.put("\u03a4", "%TAU");
		syms.put("\u03a5", "%UPSILON");
		syms.put("\u03a6", "%PHI");
		syms.put("\u03a7", "%XI");
		syms.put("\u03a8", "%PSI");
		syms.put("\u03a9", "%OMEGA");
		
		// small greek letters
		syms.put("\u03b1", "%alpha");
		syms.put("\u03b2", "%beta");
		syms.put("\u03b3", "%gamma");
		syms.put("\u03b4", "%delta");
		syms.put("\u03b5", "%varepsilon"); // epsilon, actually (Sun's idiocity)
		syms.put("\u03B6", "%zeta");
		syms.put("\u03b7", "%eta");
		syms.put("\u03b8", "%theta");
		syms.put("\u03b9", "%iota");
		syms.put("\u03ba", "%kappa");
		syms.put("\u03bb", "%lambda");
		syms.put("\u03bc", "%mu");
		syms.put("\u03bd", "%nu");
		syms.put("\u03be", "%xi");
		syms.put("\u03bf", "%omicron");
		syms.put("\u03c0", "%pi");
		syms.put("\u03c1", "%rho");
		syms.put("\u03c2", "%varsigma"); // a.k.a "final sigma" or "stigma"
		syms.put("\u03c3", "%sigma");
		syms.put("\u03c4", "%tau");
		syms.put("\u03c5", "%upsilon");
		syms.put("\u03c6", "%varphi");  // phi, actually (Sun's idiocity)
		syms.put("\u03c7", "%xi");
		syms.put("\u03c8", "%psi");
		syms.put("\u03c9", "%omega");
		syms.put("\u03d1", "%vartheta"); // a.k.a. "theta symbol"
		syms.put("\u03d5", "%phi"); // a.k.a "phi symbol"
		syms.put("\u03d6", "%varpi"); // a.k.a "pi symbol" or "omega pi"

		// special symbols
		syms.put("\u221e", "infinity");
		syms.put("\u2205", "emptyset");
		syms.put("\u2102", "setc");
		syms.put("\u2115", "setn");
		syms.put("\u211a", "setq");
		syms.put("\u211d", "setr");
		syms.put("\u2124", "setz");

		// MathType sometimes mistreats specials glyphs in text as fake equations.
		fakes.put("\u2013", "EN-DASH");
		fakes.put("\u2116", "NUMERO-SIGN");
		fakes.put("\u2044", "FRACTION-SLASH");
		fakes.put("\u2248", "ALMOST-EQUAL");
		fakes.put("forall", "FORALL");
		fakes.put("\u00bd", "FRACION-ONE-HALF");
		fakes.put("\ufffd", "PLACEBO-GLYPH");
		fakes.put("in", "IN");
		fakes.put("` in `", "IN");
		fakes.put("notin", "NOTIN");
		fakes.put("` notin `", "NOTIN");
		fakes.put("exists", "EXISTS");
		fakes.put("", "NOTHING");
		fakes.put("\u2329", "BIG_LEFT_ANGLE");
		fakes.put("\u232a", "BIG_RIGHT_ANGLE");
		fakes.put("\u2205", "EMPTYSET");
		fakes.put("emptyset", "EMPTYSET");
		fakes.put("\u227a", "PRECEDES");
		fakes.put("\u227b", "SUCCEDES");
		
		translations_ok = true;
	}

    /*
     * ========================= Annotation ===============================
     */

	private void blank () {
		switch (anno_last) {
		case 0:
		case ' ':
		case '_':
		case '^':
		case '{':
		case '(':
		case '[':
			break;
		default:
			buf.append(' ');
			anno_last = ' ';
			break;
		}
	}
	
	private void annotate(char c) {
		buf.append(c);
		anno_last = c;
	}
	
	private void annotate (String s) {
		int m = s.length();
		if (m == 0)
			return;
		buf.append(s);
		anno_last = s.charAt(m-1);
	}
	
	private static boolean isMergebleText (String s) {
		return isLatinString(s);
	}

	private void annotateLatin (String s) {
		int m = s.length();
		if (m == 0)
			return;
		if (isLatinLetter(anno_last) && isLatinLetter(s.charAt(0))) {
			int pos = s.indexOf(' ');
			String key = pos < 0 ? s : s.substring(0, pos);
			if (syms.containsKey("%"+key))
				blank();
			else
				annotate('%');
		}
		annotate(s);
	}

	private boolean isTheElement (Node node, String name) {
		return (node != null && node.getNodeType() == Node.ELEMENT_NODE
				&& name.equals(node.getLocalName()));
	}
	
	private boolean isInnerSubOrSupNode (Node node) {
		if (node == null)
			return false;
		Element parent = (Element) node.getParentNode();
		String ln = parent.getLocalName();
		if ("msub".equals(ln) || "msup".equals(ln) || "msubsup".equals(ln)) {
			return true;
		}
		return false;
	}
	
	private boolean isTallNode (Node node) {
		if (node == null || node.getNodeType() != Node.ELEMENT_NODE)
			return false;
		String ln = node.getLocalName();
		if ("mfrac".equals(ln) || "munder".equals(ln) || "munderover".equals(ln))
			return true;
		if ("mover".equals(ln)) {
			String is_accent = ((Element)node).getAttribute("math:accent");
			if (is_accent == null || !"true".equals(is_accent))
				return true;
		}
		NodeList children = node.getChildNodes();
		int n = children.getLength();
		for (int i = 0; i < n; i++) {
			if (isTallNode(children.item(i)))
				return true;
		}
		return false;
	}

	private boolean isOpNode (Node node) {
		return isTheElement(node, "mo");
	}
	
	private boolean isIdNode (Node node) {
		return isTheElement(node, "mi");
	}
	
	private boolean isNumNode (Node node) {
		return isTheElement(node, "mn");
	}

	private boolean isBracketNode (Node node) {
		if (isTheElement(node, "mfenced")) {
			String open = ((Element)node).getAttribute("math:open");
			String close = ((Element)node).getAttribute("math:close");
			String pair = open + close;
			return brackets_glyphs.contains(pair);
		}
		if (isOpNode(node)) {
			String s = node.getTextContent();
			return s != null && s.length() == 1 && brackets.containsKey(s);
		}
		return false;
	}

	private boolean isPipeNode (Node node) {
		if (isOpNode(node)) {
			String s = node.getTextContent();
			return s != null && "|".equals(s);
		}
		return false;
	}

	private boolean isDelimNode (Node node) {
		if (isOpNode(node)) {
			String s = node.getTextContent();
			return s != null && delim_glyphs.contains(s);
		}
		return false;
	}

	private boolean isSignedNumber(Node node) {
		if (isNumNode(node)) {
			String num = node.getTextContent();
			if (num != null) {
				if (!(num.startsWith("+") || num.startsWith("-"))) {
					Node op = node.getPreviousSibling();
					if (isOpNode(op)) {
						String sign = op.getTextContent();
						if (sign != null) {
							sign = stringTranslate(sign, ops);
							if ("+".equals(sign) || "-".equals(sign)) {
								Node prev = op.getPreviousSibling();
								if (prev == null || isOpNode(prev)) {
									return true;
								}
							}
						}
					}
				}
			}
		}
		return false;
	}

	private boolean needsScalableBrackets (Element elem) {
		NodeList children = elem.getChildNodes();
		int n = children.getLength();
		if (n == 1 && isTheElement(children.item(0), "mrow")) {
			children = children.item(0).getChildNodes();
			n = children.getLength();
		}
		for (int i = 0; i < n; i++) {
			Node child = children.item(i);
			if (isIdNode(child) || isNumNode(child))
				continue;
			if (isOpNode(child)) {
				if (n == 1)
					return false;
				if (i != 0 && i != n-1)
					continue;
			}
			return true;
		}
		return false;
	}

	private void annotateMath (Element elem) throws Exception {
		String ln = elem.getLocalName();
		NodeList children = elem.getChildNodes();
		int n = children.getLength();
		if ("mo".equals(ln)) {
			String s = elem.getTextContent();
			if (s.equals(MATH_SYM_FUNC))
				return;	// function application
			s = stringTranslate(s, ops);
			Node prev = elem.getPreviousSibling();
			Node next = elem.getNextSibling();
			if (("-".equals(s) || "+".equals(s)) && isSignedNumber(next)) {
				next.setTextContent(s + next.getTextContent());
				return;
			}
			boolean is_binary_op = binary_ops.contains(s);
			boolean not_bracket = !isBracketNode(elem);
			if ("/".equals(s) && (isTallNode(prev) || isTallNode(next)))
				s = "size*2\"/\"";
			boolean is_delim = isDelimNode(elem);
			boolean need_brace = isInnerSubOrSupNode(elem);
			boolean need_quote = (need_brace && (is_binary_op || prime_glyphs.contains(s)));  
			boolean need_blank;
			need_blank = (!need_brace && not_bracket
							&& !(is_delim && (isDelimNode(prev) || isIdNode(prev)
											|| isNumNode(prev) || isBracketNode(prev))
							));
			if (need_brace)
				annotate(need_quote ? '"' : '{');
			if (need_blank)
				blank();
			if (prev == null && is_binary_op && !need_quote) {
				annotate(PHANTOM_OPERAND);
				blank();
			}
			annotate(s);
			need_blank = (!need_brace && not_bracket && next != null
							&& !(is_delim && (isDelimNode(next) || isIdNode(next)
											|| isNumNode(next) || isBracketNode(next))
							));
			if (next == null && is_binary_op && !need_quote) {
				blank();
				annotate(PHANTOM_OPERAND);
			}
			if (need_blank)
				blank();
			if (need_brace)
				annotate(need_quote ? '"' : '}');
			return;
		}
		if ("mi".equals(ln)) {
			String s = elem.getTextContent();
			s = stringTranslate(s, syms);
			boolean need_guard = backops.containsKey(s.toLowerCase());
			if (need_guard) {
				annotate('%');
				annotate(s);
			} else {
				Node prev = elem.getPreviousSibling(); 
				if (prev != null && !isDelimNode(prev) && !isIdNode(prev))
					blank();
				annotateLatin(s);
			}
			return;
		}
		if ("mn".equals(ln)) {
			String s = elem.getTextContent();
			Node prev = elem.getPreviousSibling(); 
			if (prev != null && !isDelimNode(prev) && !isIdNode(prev))
				blank();
			annotateLatin(s);
			return;
		}
		if ("msub".equals(ln)) {
			Node prev = elem.getPreviousSibling();
			if (!(prev == null || isBracketNode(prev)))
				blank();
			boolean inner = isInnerSubOrSupNode(elem);
			if (inner)
				annotate('{');
			annotateChild(children.item(0));
			annotate('_');
			annotateChild(children.item(1));
			if (inner)
				annotate('}');
			return;
		}
		if ("msup".equals(ln)) {
			boolean inner = isInnerSubOrSupNode(elem);
			if (inner)
				annotate('{');
			Node prev = elem.getPreviousSibling();
			if (!(prev == null || isBracketNode(prev)))
				blank();
			annotateChild(children.item(0));
			annotate('^');
			annotateChild(children.item(1));
			if (inner)
				annotate('}');
			return;
		}
		if ("msubsup".equals(ln)) {
			boolean inner = isInnerSubOrSupNode(elem);
			if (inner)
				annotate('{');
			Node prev = elem.getPreviousSibling();
			if (!(prev == null || isBracketNode(prev)))
				blank();
			annotateChild(children.item(0));
			annotate('_');
			annotateChild(children.item(1));
			annotate('^');
			annotateChild(children.item(2));
			if (inner)
				annotate('}');
			return;
		}
		if ("mrow".equals(ln)) {
			boolean need_brace = true;
			if (need_brace
					&& n > 1
					&& isBracketNode(children.item(0))
					&& isBracketNode(children.item(n-1))
					&& !isTheElement(children.item(0), "mfenced")
					&& !isTheElement(children.item(n-1), "mfenced"))
				need_brace = false;
			if (need_brace
					&& isBracketNode(elem.getPreviousSibling())
					&& isBracketNode(elem.getNextSibling()))
				need_brace = false;
			if (need_brace) {
				Element parent = (Element) elem.getParentNode();
				String pn = parent.getLocalName();
				if ("semantics".equals(pn))
					need_brace = false;
				else if ("mfenced".equals(pn)
							&& parent.getFirstChild() == elem
							&& parent.getLastChild() == elem
							&& isBracketNode(parent))
					need_brace = false;
			}
			String sizing = elem.getAttribute(SIZING_ATTR);
			if (sizing != null && !"".equals(sizing)) {
				blank();
				sizing = sizing.trim();
				if (sizing.startsWith("+"))
					sizing = sizing.substring(1); // Java doesn't understand "+1" :)
				int size = Integer.parseInt(sizing);
				size = SIZING_BASE + SIZING_FACTOR * size;
				annotate("size "+size);
				need_brace = true;
			}
			if (n == 2 && isSignedNumber(children.item(1))) {
				annotateChild(children.item(0));
				annotateChild(children.item(1));
				return;
			}
			if (need_brace)
				annotate("{");
			if (n == 2
					&& isBracketNode(children.item(0))
					&& children.item(0).getTextContent().equals("{")
					&& isTheElement(children.item(1), "mtable")) {
				// system of equations
				blank();
				annotate("left lbrace");
				blank();
				annotateChild(children.item(1));
				blank();
				annotate("right none");
				if (!isTheElement(elem.getParentNode(), "mrow"))
					blank();
		    } else if (n > 2 && isPipeNode(children.item(0))
		    			&& isPipeNode(children.item(n-1))) {
				blank();
				annotateLatin("abs");
				boolean abs_brace = !(n == 3 && isTheElement(children.item(1), "mrow"));
				if (abs_brace)
					annotate("{");
				for (int i = 1; i < n-1; i++)
					annotateChild(children.item(i));
				if (abs_brace)
					annotate("}");
			} else {
				annotateChildren(elem);				
			}
			if (need_brace)
				annotate("}");
			return;
		}
		if ("mfrac".equals(ln)) {
			boolean prev_left_aligned = left_aligned;
			left_aligned = false;
			boolean need_brace = prev_left_aligned;
			if (!need_brace) {
				if (elem.getPreviousSibling() != null || elem.getNextSibling() != null) {
					if (!(n == 1 && children.item(0).getLocalName().equals("mrow")))
						need_brace = true;
				}
			}
			if (need_brace)
				annotate('{');
			if (prev_left_aligned) {
				annotate("alignc");
				blank();
			}
			annotateChild(children.item(0));
			blank();
			annotate("over");
			blank();
			annotateChild(children.item(1));
			if (need_brace)
				annotate('}');
			left_aligned = prev_left_aligned;
			return;
		}
		if ("mover".equals(ln)) {
			// handle accents
			String is_accent = elem.getAttribute("math:accent");
			if (is_accent != null && "true".equals(is_accent) && n == 2) {
				String accent = children.item(1).getTextContent();
				accent = stringTranslate(accent, modifiers);
				annotate("{");
				annotate(accent);
				blank();
				annotateChild(children.item(0));
				annotate("}");
				return;
			}
		}
		if ("munderover".equals(ln) || "munder".equals(ln) || "mover".equals(ln)) {
			Element func = (Element) children.item(0);
			String fname = func.getLocalName();
			String op = null;
			String summing = null;
			if ("mi".equals(fname) || "mo".equals(fname)) {
				op = stringTranslate(func.getTextContent(), ops);
				summing = (String) summings.get(op);
			}
			boolean both = "munderover".equals(ln);
			Element sub, sup;
			sub = (both || "munder".equals(ln)) ? (Element) children.item(1) : null;
			sup = (both || "mover".equals(ln)) ? (Element) children.item(both ? 2 : 1) : null;
			int expr_idx = both ? 3 : 2;
			boolean need_brace = (n > expr_idx);
			if (summing != null) {
				blank();
				annotate(summing);
				if (sub != null) {
					blank();
					annotate("from");
					annotateChild(sub);
					blank();
				}
				if (sup != null) {
					blank();
					annotate("to");
					annotateChild(sup);
					blank();
				}
			} else {
				annotateChild(func);
				if (sub != null) {
					blank();
					annotate("csub");
					annotateChild(sub);
					blank();
				}
				if (sup != null) {
					blank();
					annotate("csup");
					annotateChild(sup);
					blank();
				}
			}
			if (need_brace)
				annotate('{');
			for (int i = expr_idx; i < n; i++)
				annotateChild(children.item(i));					
			if (need_brace)
				annotate('}');
			return;
		}
		if ("mfenced".equals(ln)) {
			String open = elem.getAttribute("math:open");
			String close = elem.getAttribute("math:close");
			open = "".equals(open) ? "left none" : stringTranslate(open, brackets);
			close = "".equals(close) ? "right none" : stringTranslate(close, brackets);
			String pair = open + close;
			boolean need_brace = !brackets_glyphs.contains(pair);
			if (!need_brace && needsScalableBrackets(elem)) {
				open = stringTranslate(open, scalable);
				close = stringTranslate(close, scalable);
				need_brace = true;
			}
			if (need_brace && elem.getPreviousSibling() != null)
				blank();
			annotateLatin(open);
			if (need_brace)
				annotate('{');
			annotateChildren(elem);
			if (need_brace)
				annotate('}');
			annotateLatin(close);
			if (need_brace && elem.getNextSibling() != null) {
				char ending = close.charAt(close.length() - 1);
				if (brackets_glyphs.indexOf(ending) < 0)
					blank();
			}
			return;
		}
		if ("mtable".equals(ln)) {
			boolean prev_left_aligned = left_aligned;
			String alignment = elem.getAttribute("math:columnalign");
			if (alignment != null && "left".equals(alignment)) {
				left_aligned = true;
				blank();
				annotateLatin("alignl");
			}
			blank();
			// use "stack" if every row has only one column
			boolean prev_single_column = single_column;
			single_column = true;
			for (int i = 0; i < n; i++) {
				Node child = children.item(i);
				if (isTheElement(child, "mtr")) {
					NodeList cols = child.getChildNodes();
					if (cols.getLength() == 1
							&& isTheElement(cols.item(0), "mtd"))
						continue;
				}
				// one of conditions is broken
				single_column = false;
				break;
			}
			if (single_column) {
				annotateLatin("stack{ ");
				annotateChildren(elem);
				blank();
				annotate("} ");
			} else {
				annotateLatin("matrix{\n");
				annotateChildren(elem);
				blank();
				annotate("}\n");
			}
			single_column = prev_single_column;
			left_aligned = prev_left_aligned;
			return;
		}
		if ("mtr".equals(ln)) {
			if (elem != elem.getParentNode().getFirstChild()) {
				blank();
				annotate(single_column ? "#" : "\n##");
				blank();
			}
			annotateChildren(elem);
			return;
		}
		if ("mtd".equals(ln)) {
			if (elem != elem.getParentNode().getFirstChild()) {
				blank();
				annotate("#");
				blank();
			}
			annotateChildren(elem);
			return;
		}
		if ("msqrt".equals(ln)) {
			blank();
			annotateLatin("sqrt");
			blank();
			annotateChildren(elem);
			return;
		}
		if ("mroot".equals(ln)) {
			blank();
			annotateLatin("nroot");
			blank();
			annotateChild(children.item(n-1));
			blank();
			boolean need_brace = n > 2;
			if (need_brace)
				annotate('{');
			for (int i = 0; i < n-1; i++)
				annotateChild(children.item(i));					
			if (need_brace)
				annotate('}');
			return;
		}
		annotateChildren(elem);
	}

	private void annotateChildren (Element elem) throws Exception {
		NodeList children = elem.getChildNodes();
		int n = children.getLength();
		for (int i = 0; i < n; i++)
			annotateChild(children.item(i)); 
	}
	
	private void annotateChild (Node child) throws Exception {
		switch(child.getNodeType()) {
		case Node.ELEMENT_NODE:
			annotateMath ((Element)child);
			break;
		case Node.TEXT_NODE:
			blank();
			annotate(stringTranslate(child.getNodeValue(), ops));
			break;
		}
	}
	
    /*
     * ========================= Trivial Annotation ===============================
     */
	
	private static String trivial_glyphs = "0123456789"
											+"abcdefghijklmnopqrstuvwxyz"
											+"BACDEFGHIJKLMNOPQRSTUVWXYZ"
											+".,";
	
	public String getTrivialAnnotation() {
		String s = annotation;
		StringBuffer b = new StringBuffer(s);
		int n = s.length();
		boolean had_sub_sup = false;
		for (int i = 0; i < n; i++) {
			char c = b.charAt(i);
			if (trivial_glyphs.indexOf(c) >= 0)
				continue;
			if (c == (char)0)
				continue; // already removed
			if (ADVANCED_TRIVIAL_SPACING && c == ' ' && i > 0 && i < n-1) {
				if (isLatinLetter(b.charAt(i+1)) && Character.isDigit(b.charAt(i-1))) {
					b.setCharAt(i, (char)0);
					continue;
				}
			}
			if (ADVANCED_TRIVIAL_SUBSCRIPTING && c == '_' || c == '^' && !had_sub_sup) {
				had_sub_sup = true;
				if (i < n-1 && b.charAt(i+1) == '{') {
					if (b.charAt(n-1) != '}')
						return null;
					b.setCharAt(i+1, (char)0);
					b.setCharAt(n-1, (char)0);
				}
				continue;
			}
			return null;
		}
		for (int i = n-1; i >= 0; i--) {
			if (b.charAt(i) == (char)0) // remove characters marked for deletion
				b.deleteCharAt(i);
		}
		String result = b.toString();
		if (ops.containsValue(result)
				|| brackets.containsValue(result)
				|| syms.containsValue(result)
				|| modifiers.containsValue(result))
			return null;
		return result;
	}

    /*
     * ========================= Refactoring from MathType ===============================
     */

	private Element addFence (Element to, String open_br, String close_br) {
		Element elem = mml_doc.createElementNS(MATH_NS, "math:mfenced");
		to.appendChild(elem);
		elem.setAttributeNS(MATH_NS, "math:open", open_br);
		elem.setAttributeNS(MATH_NS, "math:close", close_br);
		return elem;
	}

	private Element addOp (Element to, String op, Element attr_src) {
		Element elem = mml_doc.createElementNS(MATH_NS, "math:mo");
		to.appendChild(elem);
		elem.setAttributeNS(MATH_NS, "math:stretchy", "false");
		if (attr_src != null)
			copyMathAttr(attr_src, elem);
		elem.setTextContent(op);
		return elem;
	}

	private void transformMath (Element from, Element to) throws Exception {
		String ln = from.getLocalName();
		Element elem;
		if ("mi".equals(ln)) {
			// merge similar adjacent nodes
			Node next = from.getNextSibling();
			if (next != null && next.getNodeType() == Node.ELEMENT_NODE) {
				String next_ln = next.getLocalName();
				if ("mo".equals(next_ln) && next.getTextContent().equals(MATH_SYM_FUNC)) {
					// Check for MathType bug (function cannot be last in the expression)
					Element dad = (Element) from.getParentNode();
					if (dad.getLocalName().equals("mrow")
							&& dad.getChildNodes().getLength() == 2) {
						Element granddad = (Element) dad.getParentNode();
						if (dad == granddad.getLastChild()) {
							// Actually, this is THE identifier.
							to.appendChild(elem = mml_doc.createElementNS(MATH_NS, "math:mi"));
							copyMathAttr(from, elem);
							elem.setTextContent(stringTranslate(from.getTextContent(), backops));
							return;
						}
					}
					// this "ID" is a function name.
					next.setTextContent(from.getTextContent());
					return;
				} if (next_ln.equals(ln)) {
					if ("mrow".equals(from.getParentNode().getLocalName())) {
						String text1 = from.getTextContent();
						String text2 = next.getTextContent();
						if (isMergebleText(text1) && isMergebleText(text2)) {
							next.setTextContent(text1 + text2);
							return;
						}
					}
				} else if (ADVANCED_ID_MERGING) {
					if (next_ln.equals("msub") || next_ln.equals("msup")) {
						// Handle MathType bug.
						// Dangerous !
						Node sub_first = next.getFirstChild();
						if (sub_first.getLocalName().equals(ln)) {
							String text1 = from.getTextContent();
							String text2 = sub_first.getTextContent();
							if (isMergebleText(text1) && isMergebleText(text2)) {
								sub_first.setTextContent(text1 + text2);
								return;
							}
						}
					}
				}
			}
			to.appendChild(elem = mml_doc.createElementNS(MATH_NS, "math:mi"));
			copyMathAttr(from, elem);
			elem.setTextContent(stringTranslate(from.getTextContent(), backops));
			return;
		}
		if ("mo".equals(ln)) {
			String op = from.getTextContent();
			if (op.equals(MATH_SYM_FUNC))
				return;
			if (":".equals(op)) {
				// sense ":="
				Node next = from.getNextSibling();
				if (next != null && next.getNodeType() == Node.ELEMENT_NODE
						&& "mo".equals(next.getLocalName())) {
					String next_op = next.getTextContent();
					if (next_op != null && "=".equals(next_op)) {
						next.setTextContent(MATH_ASSIGNMENT);
						return;
					}
				}
			}
			addOp(to, stringTranslate(op, backops), from);
			return;
		}
		if ("mn".equals(ln)) {
			String num = from.getTextContent();
			// MathType glitch: always associate first dots with the following number
			Node prev = from.getPreviousSibling();
			if (num.startsWith(".")	&&
					(isNumNode(prev) || isIdNode(prev) || isBracketNode(prev))) {
				int pos = 1;
				int len = num.length();
				while (pos < len && num.charAt(pos) == '.')
					pos++;
				addOp(to, num.substring(0, pos), null);
				num = num.substring(pos);
			}
			// MathType glitch: always associate last dots with the previous number
			String dots = null;
			Node next = from.getNextSibling();
			if (num.endsWith(".") && (isIdNode(next))) {
				int pos = num.length() - 1;
				while (pos > 0 && num.charAt(pos-1) == '.')
					pos--;
				dots = num.substring(pos);
				num = num.substring(0, pos);
			}
			to.appendChild(elem = mml_doc.createElementNS(MATH_NS, "math:mn"));
			copyMathAttr(from, elem);
			elem.setTextContent(num);
			if (dots != null)
				addOp(to, dots, null);
			return;
		}
		if ("mstyle".equals(ln) || "mrow".equals(ln)) {
			NodeList children = from.getChildNodes();
			int n = children.getLength();
			if ("mstyle".equals(ln)) {
				String sizing = from.getAttribute(SIZING_ATTR);
				if (sizing != null && !"".equals(sizing)) {
					// keep the element to keep the size information
					to.appendChild(elem = mml_doc.createElementNS(MATH_NS, "math:mrow"));
					elem.setAttribute(SIZING_ATTR, sizing);
					transformChildren (from, elem);
					return;
				}
			}
			if (n == 1 && children.item(0).getNodeType() == Node.ELEMENT_NODE) {
				String cln = children.item(0).getLocalName();
				if ("mstyle".equals(cln) || "mrow".equals(cln)
						|| "mfrac".equals(cln) || "msub".equals(cln)
						|| "mtable".equals(cln)) {
					transformChildren (from, to);
					return;
				}
			}
			if (n >= 2) {
				Node first = children.item(0);
				Node last = children.item(n-1);
				if (n == 2 && isOpNode(last) && last.getTextContent().equals(MATH_SYM_FUNC)) {
					// function application...
					transformChildren (from, to);
					return;
				}
				if (n == 2 && isBracketNode(first) && first.getTextContent().equals("{")) {
					if (isTheElement(last, "mtable")) {
						// System of equations
						transformChild(last, addFence(to, "{", ""));
						return;
					}
					if (isTheElement(last, "mrow")
							&& last.getChildNodes().getLength() == 1) {
						Node sub_child = last.getFirstChild();
						if (isTheElement(sub_child, "mtable")) {
							// System of equations
							transformChild(sub_child, addFence(to, "{", ""));
							return;
						}
					}
				}
				if (n >= 2 && isBracketNode(first) && isBracketNode(last)) {
					// expression in brackets
					elem = addFence(to, stringTranslate(first.getTextContent(), brackets),
									stringTranslate(last.getTextContent(), brackets));
					if (n == 3 && isTheElement(children.item(1), "mrow")) {
						// avoid excess braces
						transformChildren((Element)children.item(1), elem);
						return;
					}
					for (int i = 1; i < n - 1; i++)
						transformChild(children.item(i), elem);
					return;
				}
			}
			to.appendChild(elem = mml_doc.createElementNS(MATH_NS, "math:mrow"));
			transformChildren (from, elem);
			return;
		}
		if ("semantics".equals(ln)) {
			// skip attributes
			to.appendChild(elem = mml_doc.createElementNS(MATH_NS, "math:semantics"));
			transformChildren (from, elem);
			return;
		}
		if ("annotation".equals(ln)) {
			return;
		}
		// all other elements are simply copied
		to.appendChild(elem = mml_doc.createElementNS(MATH_NS, "math:"+from.getLocalName()));
		copyMathAttr(from, elem);
		transformChildren (from, elem);
	}

	private void transformChildren (Element from, Element to) throws Exception {
		NodeList children = from.getChildNodes();
		int n = children.getLength();
		for (int i = 0; i < n; i++) {
			transformChild(children.item(i), to); 
		}
	}
	
	private void transformChild (Node child, Element to) throws Exception {
		switch(child.getNodeType()) {
		case Node.ELEMENT_NODE:
			transformMath ((Element)child, to);
			break;
		default:
			to.appendChild(mml_doc.importNode(child, true));
			break;
		}
	}

	private static void copyMathAttr (Element from, Element to) {
		NamedNodeMap atts = from.getAttributes();
		int n = atts.getLength();
		for (int i = 0; i < n; i++) {
			Attr a = (Attr) atts.item(i);
			String pfx = a.getPrefix();
			String ln = a.getLocalName();
			if (pfx != null && "xmlns".equals(pfx))
				continue;
			if ("xmlns".equals(ln))
				continue;
			to.setAttributeNS(MATH_NS, "math:"+ln, a.getNodeValue());
		}
	}

    /*
     * ============================= Utilities ============================
     */

	private static final String stringTranslate (String str, HashMap translations) {
		String result = (String) translations.get(str);
		return (result != null ? result : str);
	}
	
	private static final void reverseHashMap (HashMap from, HashMap to) {
		Iterator it = from.entrySet().iterator();
		while(it.hasNext()) {
			Map.Entry entry = (Map.Entry) it.next();
			to.put(entry.getValue(), entry.getKey());
		}
	}
	
	private static final boolean isLatinLetter (char c) {
		return ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'));
	}
	
	private static boolean isLatinString (String s) {
		int n = s.length();
		for (int i = 0; i < n; i++) {
			if (!isLatinLetter(s.charAt(i)))
				return false;
		}
		return true;
	}
	
}
