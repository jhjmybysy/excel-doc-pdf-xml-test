package com.test.xml;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Ϊ���DOM�����⣬������SAX��SAX
 * ���¼�������������������Ԫ�ؿ�ʼ��Ԫ�ؽ������ı����ĵ��Ŀ�ʼ�������ʱ�������¼�������Ա��д��Ӧ��Щ�¼��Ĵ��룬��������
 * ���ŵ㣺�������ȵ��������ĵ���ռ����Դ��
 * ��SAX�����������DOM����������С������Applet�����ء�ȱ�㣺���ǳ־õģ��¼�������û�������ݣ���ô���ݾͶ���
 * ����״̬�ԣ����¼���ֻ�ܵõ��ı�������֪���ı������ĸ�Ԫ�أ�ʹ�ó��ϣ�Applet;ֻ��XML�ĵ����������ݣ����ٻ�ͷ���ʣ������ڴ��٣�
 */
public class SaxDemo implements XmlDocument {

	public static void main(String[] args) {
		SaxDemo demo = new SaxDemo();
		demo.parserXml("text.xml");
	}

	public void createXml(String fileName) {
		System.out.println("<<" + fileName + ">>");
	}

	public void parserXml(String fileName) {
		SAXParserFactory saxfac = SAXParserFactory.newInstance();
		try {
			SAXParser saxparser = saxfac.newSAXParser();
			InputStream is = new FileInputStream(fileName);
			saxparser.parse(is, new MySAXHandler());
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		} catch (SAXException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

class MySAXHandler extends DefaultHandler {
	boolean hasAttribute = false;
	Attributes attributes = null;

	public void startDocument() throws SAXException {
		System.out.println("�ĵ���ʼ��ӡ��");
	}

	public void endDocument() throws SAXException {
		System.out.println("�ĵ���ӡ������");
	}

	public void startElement(String uri, String localName, String qName,
			Attributes attributes) throws SAXException {
		if (qName.equals("employees")) {
			return;
		}
		if (qName.equals("employee")) {
			System.out.println(qName);
		}
		if (attributes.getLength() > 0) {
			this.attributes = attributes;
			this.hasAttribute = true;
		}
	}

	public void endElement(String uri, String localName, String qName)
			throws SAXException {
		if (hasAttribute && (attributes != null)) {
			for (int i = 0; i < attributes.getLength(); i++) {
				System.out.println(attributes.getQName(0)
						+ attributes.getValue(0));
			}
		}
	}

	public void characters(char[] ch, int start, int length)
			throws SAXException {
		System.out.println(new String(ch, start, length));
	}
}
