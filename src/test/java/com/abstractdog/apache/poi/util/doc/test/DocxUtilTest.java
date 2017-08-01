package com.abstractdog.apache.poi.util.doc.test;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Assert;
import org.junit.Test;

import com.abstractdog.apache.poi.util.doc.DocxUtil;

public class DocxUtilTest {
	
	@Test
	public void testReadDocx() throws Exception{
		XWPFDocument doc = DocxUtil.readDocxFile(new File("src/main/resources/doc/helloworld.docx"));
		Assert.assertNotEquals(0, doc.getParagraphs().size());
	}
	
	@Test
	public void testChangeText() throws Exception{
		XWPFDocument doc = DocxUtil.changeText(new File("src/main/resources/doc/helloworld.docx"), "world", "*****");
		doc.write(new FileOutputStream(new File("src/test/resources/out/helloworld_out.docx")));
	}
	
	@Test
	public void testChangeRegex() throws Exception{
		XWPFDocument doc = DocxUtil.changeRegex(new File("src/main/resources/doc/helloworld.docx"), "[a-z]", "*");
		doc.write(new FileOutputStream(new File("src/test/resources/out/helloworld_out.docx")));
	}
}
