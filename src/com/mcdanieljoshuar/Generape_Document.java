package com.mcdanieljoshuar;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

public class Generape_Document {

	public static void main(String[] args) { 
		
		try {
			
			//This creates a blank document.
			XWPFDocument document = new XWPFDocument();
			//This is the path the document will be saved to.
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Intern\\Desktop\\PAPA_Docx\\WordDocuments\\APAWord.docx"));
									
			CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
			XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
			
			//write header content
			CTP ctpHeader = CTP.Factory.newInstance();
		    CTR ctrHeader = ctpHeader.addNewR();
			CTText ctHeader = ctrHeader.addNewT();
			String headerText = "Running Head: SHORT TITLE";
			ctHeader.setStringValue(headerText);
			XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
		    XWPFParagraph[] parsHeader = new XWPFParagraph[1];
		    parsHeader[0] = headerParagraph;
		    policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);
			
			//Creates a title.
			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run = paragraph.createRun();
			paragraph.setAlignment(ParagraphAlignment.CENTER);
			paragraph.setSpacingBetween(2);
			run.setFontFamily("TIMES NEW ROMAN");
			run.setFontSize(12);
			run.addBreak();
			run.addBreak();
			run.addBreak();
			run.addBreak();
			run.addBreak();
			run.addBreak();
			run.addBreak();
			run.addBreak();
			run.setText("Title");
			run.addBreak();
			run.setText("Name");
			paragraph.setPageBreak(true);
			document.write(out);
			out.close();
			System.out.println("Word doc printed successfully!");
			
			} catch (Exception e) {
				System.out.println("Failed creating a paragraph!");
			}

	}

}
