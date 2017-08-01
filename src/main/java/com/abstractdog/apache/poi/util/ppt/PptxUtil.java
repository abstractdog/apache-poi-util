package com.abstractdog.apache.poi.util.ppt;

import java.io.File;
import java.io.FileInputStream;
import java.util.function.Consumer;
import java.util.regex.Pattern;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PptxUtil {
	public static XMLSlideShow readPptxFile(File file) throws Exception {
		FileInputStream fis = new FileInputStream(file.getAbsolutePath());

		XMLSlideShow ppt = new XMLSlideShow(fis);
		fis.close();
		
		return ppt;
	}

	public static XMLSlideShow changeText(File file, String from, String to) throws Exception {
		XMLSlideShow ppt = readPptxFile(file);

		Consumer<XSLFTextRun> textChanger = (run) -> {
			String text = run.getText();

			if (text != null && text.contains(from)) {
				text = text.replace(from, to);
				run.setText(text);
			}
		};
		
		applyRunConsumer(ppt, textChanger);
		
		return ppt;
	}

	public static XMLSlideShow changeRegex(File file, String from, String to) throws Exception {
		XMLSlideShow ppt = readPptxFile(file);

		Consumer<XSLFTextRun> regexChanger = (run) -> {
			String text = run.getText();

			if (text != null) {
				run.setText(Pattern.compile(from).matcher(text).replaceAll(to));
			}
		};
		
		applyRunConsumer(ppt, regexChanger);
		
		return ppt;
	}
	
	private static void applyRunConsumer(XMLSlideShow ppt, Consumer<XSLFTextRun> runConsumer) {
		XSLFSlide[] slides = ppt.getSlides();
		
		for (int i = 0; i < slides.length; i++) {
			for (XSLFShape shape : slides[i].getShapes()) {
				if (shape instanceof XSLFTextShape) {
					XSLFTextShape txShape = (XSLFTextShape) shape;
					for (XSLFTextParagraph xslfParagraph : txShape.getTextParagraphs()) {
						for (XSLFTextRun run : xslfParagraph.getTextRuns()){
							runConsumer.accept(run);
						}
					}
				}
			}
		}
	}
}
