package com.test.xml;

/**
 * ����XML�ĵ�����������Ľӿ�
 */
public interface XmlDocument {
	/**
	 * ����XML�ĵ�
	 * 
	 * @param fileName
	 *            �ļ�ȫ·������
	 */
	public void createXml(String fileName);

	/**
	 * ����XML�ĵ�
	 * 
	 * @param fileName
	 *            �ļ�ȫ·������
	 */
	public void parserXml(String fileName);
}
