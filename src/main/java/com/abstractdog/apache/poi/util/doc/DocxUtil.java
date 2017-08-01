package com.abstractdog.apache.poi.util.doc;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;
import java.util.function.Consumer;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class DocxUtil {
	public static XWPFDocument readDocxFile(File file) throws Exception {
		FileInputStream fis = new FileInputStream(file.getAbsolutePath());

		XWPFDocument document = new XWPFDocument(fis);
		fis.close();

		return document;
	}
	
	public static XWPFDocument changeText(File file, String from, String to) throws Exception {
		XWPFDocument doc = readDocxFile(file);
		
		Consumer<XWPFRun> textChanger = (run) -> {
			String text = run.getText(0);

			if (text != null && text.contains(from)) {
				text = text.replace(from, to);
				run.setText(text, 0);
			}
		};
		
		applyRunConsumer(doc, textChanger);
		
		return doc;
	}

	public static XWPFDocument changeRegex(File file, String from, String to) throws Exception {
		XWPFDocument doc = readDocxFile(file);
		
		Consumer<XWPFRun> regexChanger = (run) -> {
			String text = run.getText(0);

			if (text != null) {
				run.setText(Pattern.compile(from).matcher(text).replaceAll(to), 0);
			}
		};
		
		applyRunConsumer(doc, regexChanger);
		
		return doc;
	}
	
	private static void applyRunConsumer(XWPFDocument doc, Consumer<XWPFRun> runConsumer) {
		for (XWPFParagraph p : doc.getParagraphs()) {
			List<XWPFRun> runs = p.getRuns();
			if (runs != null) {
				for (XWPFRun run : runs) {
					runConsumer.accept(run);
				}
			}
		}
		
		for (XWPFTable tbl : doc.getTables()) {
			for (XWPFTableRow row : tbl.getRows()) {
				for (XWPFTableCell cell : row.getTableCells()) {
					for (XWPFParagraph p : cell.getParagraphs()) {
						for (XWPFRun run : p.getRuns()) {
							runConsumer.accept(run);
						}
					}
				}
			}
		}
	}
}
