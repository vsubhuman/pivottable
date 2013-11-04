package com.vsubhuman.xml;

import java.util.Iterator;

import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * <p>Class provides interfaces of {@link Iterator} and {@link Iterable} only
 * for {@link Element} nodes from specified {@link NodeList}.</p>
 * 
 * <p>Example of use:
 * <pre>
 * Document doc = ...
 * Element root = doc.getDocumentElement();
 * ElementIterator ei = ElementIterator.create(root, "element");
 * for (Element e : ei) {
 *   ...
 * }</pre>
 * 
 * <p>For every not-element node special method {@link #notElement(NodeList, int, Node)}
 * called. You can implements some functionality by extending class and overriding
 * this method.
 * 
 * <p><b>Note:</b> you shoud not make changes to iterated node list or parent
 * element (such as delete elements from it) until iteration is ended,
 * cuz it will corrupt order of elements. 
 * 
 * @author vsubhuman
 * @version 1.0
 */
public class ElementIterator implements Iterator<Element>, Iterable<Element> {

	// List of nodes
	private NodeList nodes;
	// Current index
	private int index;

	/**
	 * Creates an instance of the iterator for specified node list. 
	 * @param nodes - {@link NodeList} to iterate
	 * @since 1.0
	 */
	public ElementIterator(NodeList nodes) {

		this.nodes = nodes;
	}
	
	@Override
	public Iterator<Element> iterator() {
		return this;
	}

	/**
	 * Returns <code>true</code> if node list has eny more element nodes in it.
	 */
	@Override
	public boolean hasNext() {
		
		while (index < nodes.getLength()) {
			
			Node n = nodes.item(index);
			if (n.getNodeType() == Node.ELEMENT_NODE) {

				return true;
			}
			else {

				notElement(nodes, index, n);
				index++;
			}
		}
		
		return false;
	}
	
	/**
	 * This method called for every not-element node in the iteration. 
	 * @param nodes - {@link NodeList} used for iteration
	 * @param index - index of the current node
	 * @param node - current node
	 * @since 1.0
	 */
	protected void notElement(NodeList nodes, int index, Node node) {};
	
	/**
	 * @throws IllegalStateException is no more elements available
	 */
	@Override
	public Element next() throws IllegalStateException {
		
		if (!hasNext())
			throw new IllegalStateException("No more elements available!");
		
		return (Element) nodes.item(index++); 
	}
	
	/**
	 * Not supported.
	 * @throws UnsupportedOperationException in any case.
	 */
	public void remove() throws UnsupportedOperationException {
		throw new UnsupportedOperationException();
	}
	
	/**
	 * Creates new instance of the iterator with {@link NodeList} geted
	 * from specified element by specified tag name.
	 * @param parent - element to iterate nodes of.
	 * @param elementName - tag name of iterated elements
	 * @return new {@link ElementIterator}
	 * @since 1.0
	 */
	public static ElementIterator create(Element parent, String elementName) {
		
		NodeList l = parent.getElementsByTagName(elementName);
		return new ElementIterator(l);
	}
}
