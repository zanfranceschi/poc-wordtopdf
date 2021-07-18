package pocs.arquitetura.poc_wordtopdf;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import lombok.SneakyThrows;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cache.annotation.Cacheable;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;


@SpringBootApplication
@RestController
public class PocWordtopdfApplication {

	public static void main(String[] args) {
		SpringApplication.run(PocWordtopdfApplication.class, args);
	}

	Logger logger = LoggerFactory.getLogger(PocWordtopdfApplication.class);


	@SneakyThrows
	@GetMapping("/teste")
	public String teste() {

		String docxOutput = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato.docx";
		String pdfOutput = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato.pdf";

		try (XWPFDocument doc = getXWPFDocument()) {

			List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
			for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
				for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {

					String docText = xwpfRun.getText(0);
					logger.info(xwpfRun.toString());
					logger.info(docText);

					if (docText != null) {
						docText = docText.replace("${NOME}", "Anna Bernardes Rezende Menezes");
						docText = docText.replace("${DOCUMENTO}", "123.123.123-12");
						docText = docText.replace("${DATA}", "18/07/2021");
						xwpfRun.setText(docText, 0);
					}
				}
			}

			try (FileOutputStream out = new FileOutputStream(docxOutput)) {
				doc.write(out);
			}

			docx2pdf(docxOutput, pdfOutput);
		}
		return "ok";
	}

	@SneakyThrows
	private void docx2pdf(String docxPath, String pdfPath) {
		InputStream in = new FileInputStream(docxPath);
		XWPFDocument document = new XWPFDocument(in);
		PdfOptions options = PdfOptions.create();
		OutputStream out = new FileOutputStream(pdfPath);
		PdfConverter.getInstance().convert(document, out, options);
		document.close();
		out.close();
	}

	@Cacheable("template")
	@SneakyThrows
	private XWPFDocument getXWPFDocument() {
		String templatePath = getClass().getResource("/contrato-template.docx").getPath();
		XWPFDocument doc = new XWPFDocument(Files.newInputStream(Path.of(templatePath)));
		return doc;
	}

}
