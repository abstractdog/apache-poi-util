package com.abstractdog.apache.poi.util.ppt.test;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.junit.Assert;
import org.junit.Test;

import com.abstractdog.apache.poi.util.ppt.PptxUtil;

public class PptxUtilTest {
	
	@Test
	public void testReadPptx() throws Exception{
		XMLSlideShow ppt = PptxUtil.readPptxFile(new File("src/main/resources/ppt/helloworld.pptx"));
		Assert.assertNotEquals(0, ppt.getSlides().length);
	}
	
	@Test
	public void testChangeText() throws Exception{
		XMLSlideShow ppt = PptxUtil.changeText(new File("src/main/resources/ppt/helloworld.pptx"), "world", "*");
		ppt.write(new FileOutputStream(new File("src/test/resources/out/helloworld_out.pptx")));
	}
	
	@Test
	public void testChangeRegex() throws Exception{
		XMLSlideShow doc = PptxUtil.changeRegex(new File("src/main/resources/ppt/helloworld.pptx"), "[a-z]", "*");
		doc.write(new FileOutputStream(new File("src/test/resources/out/helloworld_out.pptx")));
	}
}
