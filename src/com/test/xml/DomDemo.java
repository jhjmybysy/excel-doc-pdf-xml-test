package com.test.xml;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * 为 XML 文档的已解析版本定义了一组接口。解析器读入整个文档，然后构建一个驻留内存的树结构，然后代码就可以使用 DOM
 * 接口来操作这个树结构。优点：整个文档树在内存中
 * ，便于操作；支持删除、修改、重新排列等多种功能；缺点：将整个文档调入内存（包括无用的节点），浪费时间和空间；使用场合
 * ：一旦解析了文档还需多次访问这些数据；硬件资源充足（内存、CPU）。
 * 
 * DOM生成与解析XML文档
 */
public class DomDemo implements XmlDocument {
	private Document document;
	private String fileName;

	public void init() {
		try {
			DocumentBuilderFactory factory = DocumentBuilderFactory
					.newInstance();
			DocumentBuilder builder = factory.newDocumentBuilder();
			this.document = builder.newDocument();
		} catch (ParserConfigurationException e) {
			System.out.println(e.getMessage());
		}
	}

	public void createXml(String fileName) {
		Element root = this.document.createElement("employees");
		this.document.appendChild(root);
		Element employee = this.document.createElement("employee");
		Element name = this.document.createElement("name");
		name.appendChild(this.document.createTextNode("你好"));
		employee.appendChild(name);
		Element sex = this.document.createElement("sex");
		sex.appendChild(this.document.createTextNode("m"));
		employee.appendChild(sex);
		Element age = this.document.createElement("age");
		age.appendChild(this.document.createTextNode("30"));
		employee.appendChild(age);
		root.appendChild(employee);
		TransformerFactory tf = TransformerFactory.newInstance();
		try {
			Transformer transformer = tf.newTransformer();
			DOMSource source = new DOMSource(document);
			transformer.setOutputProperty(OutputKeys.ENCODING, "gb2312");
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			PrintWriter pw = new PrintWriter(new FileOutputStream(fileName));
			StreamResult result = new StreamResult(pw);
			transformer.transform(source, result);
			System.out.println("生成XML文件成功!");
		} catch (TransformerConfigurationException e) {
			System.out.println(e.getMessage());
		} catch (IllegalArgumentException e) {
			System.out.println(e.getMessage());
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		} catch (TransformerException e) {
			System.out.println(e.getMessage());
		}
	}

	public void parserXml(String fileName) {
		try {
			DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
			DocumentBuilder db = dbf.newDocumentBuilder();
			Document document = db.parse(fileName);
			NodeList employees = document.getChildNodes();
			for (int i = 0; i < employees.getLength(); i++) {
				Node employee = employees.item(i);
				NodeList employeeInfo = employee.getChildNodes();
				for (int j = 0; j < employeeInfo.getLength(); j++) {
					Node node = employeeInfo.item(j);
					NodeList employeeMeta = node.getChildNodes();
					for (int k = 0; k < employeeMeta.getLength(); k++) {
						System.out.println(employeeMeta.item(k).getNodeName()
								+ ":" + employeeMeta.item(k).getTextContent());
					}
				}
			}
			System.out.println("解析完毕");
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		} catch (ParserConfigurationException e) {
			System.out.println(e.getMessage());
		} catch (SAXException e) {
			System.out.println(e.getMessage());
		} catch (IOException e) {
			System.out.println(e.getMessage());
		}
	}
	
	
	public static void main(String[] args) {
		DomDemo demo = new DomDemo();
		demo.init();
		demo.createXml("text.xml");
		demo.parserXml("text.xml");
	}
}
