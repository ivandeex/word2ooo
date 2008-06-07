package net.vitki.ooo7;

import java.io.File;

import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * @author vit
 *
 */

public class Util
{
    /*
     * ============================= XML utilities ============================
     */

    protected static Node normalizeElement (Node node) throws Exception {
    	node.normalize();
   		normalizeSubtree(node);
    	return node;
    }

    public static final String normalizeText(String text) {
		StringBuffer sb = new StringBuffer();
		int n = text.length();
		boolean prev_blank = false;
		for (int i = 0; i < n; i++) {
			char c = text.charAt(i);
			if (c == '\n' || Character.isWhitespace(c)) {
				if (prev_blank)
					continue;
				sb.append(' ');
				prev_blank = true;
			} else {
				prev_blank = false;
				sb.append(c);
			}
		}
		return sb.toString().trim();
    }

    public static final boolean isWhiteSpaceString(String text) {
		int n = text.length();
		for (int i = 0; i < n; i++) {
			char c = text.charAt(i);
			if (c != ' ' && c != '\t' && c != '\n' && c != '\r')
				return false;
		}
		return true;
    }

    private static void normalizeSubtree(Node node) throws Exception {
    	NodeList children = node.getChildNodes();
    	int i;
    	for (i = children.getLength() - 1; i >= 0; i--) {
    		Node child = children.item(i);
    		switch (child.getNodeType()) {
    		case Node.ELEMENT_NODE:
    			normalizeSubtree(child);
    			break;
    		case Node.TEXT_NODE:
    			if (isWhiteSpaceString(child.getNodeValue()))
    				node.removeChild(child);
    			break;
    		}
    	}
    }

    /*
     * ============================= File utilities ============================
     */

	public static final String getFileName (String path) {
		File file = new File(path);
		try {
			path = file.getCanonicalPath();
		} catch (Exception e) {
			path = file.getAbsolutePath();
		}
		int dot = path.lastIndexOf('.');
		if (dot >= 0) {
			path = path.substring(0, dot);
		}
		return path;
	}
	
    public static final void cleanDirectory (String path) {
		File d = new File(path);
		if (!d.isDirectory())
			d.delete();
		d.mkdir();
		File[] lst = d.listFiles();
   		for (int i = 0; i < lst.length; i++)
   			lst[i].delete();
    }

    public static final boolean fileExists (String path) {
    		return (new File(path)).exists();
    }

    /*
     * ============================= String utilities ============================
     */

	public static final String num000(int no) {
		if (no < 10)
			return "000"+no;
		else if (no < 100)
			return "00"+no;
		else if (no < 1000)
			return "0"+no;
		else
			return ""+no;
	}

    public static final boolean isOneOf (String sample, String set, boolean ignore_case) {
    	if (sample == null || set == null)
    		return false;
    	int sample_len = sample.length();
    	int set_len = set.length();
    	if (sample_len == 0 || set_len == 0)
    		return false;
    	if (ignore_case) {
    		set = set.toLowerCase();
    		sample = sample.toLowerCase();
    	}
    	int pos = set.indexOf(sample);
    	if (pos < 0)
    		return false;
    	if (pos == 0 || set.charAt(pos-1) == ' ') {
    		int end = pos + sample_len;
    		if (end >= set_len || set.charAt(end) == ' ')
    			return true;
    	}
    	return false;
    }

    public static final boolean isOneOf (String sample, String set) {
    	return isOneOf(sample, set, false);
    }

    /*
     * ============================= Other utilities ============================
     */

	public static final String normalRef (String ref) {
		return (ref == null ? null : (ref.startsWith("./") ? ref.substring(2) : ref));
	}
	
	public static final String image_formats = "BMP EPS GIF JPG JPEG PCX PNG SVG TIFF TIF WMF";

	public static final boolean isImageFormat (String fmt) {
		return isOneOf(fmt, image_formats, true);
	}

    public static final String contents_titles = "contents contents. Содержание Содержание. СОДЕРЖАНИЕ";

	public static final boolean isContentsTitle (String title) {
		return isOneOf(title, contents_titles, true);
	}

}
