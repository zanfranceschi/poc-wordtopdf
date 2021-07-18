package pocs.arquitetura.poc_wordtopdf;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.*;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cache.annotation.Cacheable;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.xml.bind.JAXBElement;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Base64;
import java.util.List;


@SpringBootApplication
@RestController
public class PocWordtopdfApplication {

	public static void main(String[] args) {
		SpringApplication.run(PocWordtopdfApplication.class, args);
	}

	Logger logger = LoggerFactory.getLogger(PocWordtopdfApplication.class);

	public static File docxContratoTemplateFile = new File(
			PocWordtopdfApplication.class.getResource("/contrato-template.docx").getPath()
	);

	@SneakyThrows
	@GetMapping("/pdf")
	public Resposta pdf() {

		String docxOutputPath = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato2.docx";
		String pdfOutputPath = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato2.pdf";

		// Com WordprocessingMLPackage é possível busar nodes com XPath, com XWPFDocument não.
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(docxContratoTemplateFile);

		MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

		String textNodesXPath = "//w:t"; // <w:t/> é o nó folha do documento XML que contém texto.
		List<Object> textNodes = mainDocumentPart.getJAXBNodesViaXPath(textNodesXPath, true);
		for (Object obj : textNodes) {
			Text text = (Text) ((JAXBElement) obj).getValue();
			// substituição dos placeholders (qq string arbitrária)
			text.setValue(text.getValue().replace("${NOME}", "CHICO"));
			text.setValue(text.getValue().replace("${DATA}", "20-20-10"));
			text.setValue(text.getValue().replace("${DOCUMENTO}", "123.123.1231-22"));
		}

		try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
			wordMLPackage.save(out);

			// transforma em objeto XWPFDocument para converter para PDF
			try(ByteArrayInputStream in = new ByteArrayInputStream(out.toByteArray())) {

				try (XWPFDocument document = new XWPFDocument(in)) {
					// Desculpa, me perdi nos aninhamntos com os "tries"
					// com .NET dá pra deixar mais limpo ;)

					boolean debug = true;
					if (debug) { // salvar em disco para verificar o resultado

						wordMLPackage.save(new File(docxOutputPath)); // salva o docx em disco

						try (FileOutputStream pdfOutputStream = new FileOutputStream(pdfOutputPath)) { // salva o pdf em disco
							PdfOptions options = PdfOptions.create();
							PdfConverter.getInstance().convert(document, pdfOutputStream, options);
						}
					}

					try (ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream()) {

						PdfOptions options = PdfOptions.create();
						PdfConverter.getInstance().convert(document, pdfOutputStream, options);

						byte[] pdfBytes = Base64.getEncoder().encode(pdfOutputStream.toByteArray());
						String pdfBase64 = new String(pdfBytes);
						return new Resposta(pdfBase64, "teste");
					}
				}
			}
		}
	}

	@SneakyThrows
	@GetMapping("/teste")
	public Resposta teste() {
		// não use esse método... o método "pdf" tá melhor... mais inteligível e funciona com tabelas...

		String docxTemplatePath = getClass().getResource("/contrato-template.docx").getPath();

		String docxOutputPath = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato.docx";
		String pdfOutputPath = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato.pdf";

		try (XWPFDocument doc = getXWPFDocument()) {

			List<XWPFParagraph> xwpfParagraphList = doc.getParagraphs();
			for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
				for (XWPFRun xwpfRun : xwpfParagraph.getRuns()) {

					logger.info(String.format("hash code %s", xwpfRun.hashCode()));

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

			// gravar o arquivo docx em disco pra testar mais facilmente
			try (FileOutputStream docxOutputStream = new FileOutputStream(docxOutputPath)) {
				doc.write(docxOutputStream);
			}

			try (ByteArrayOutputStream docxOutputStream = new ByteArrayOutputStream()) {
				doc.write(docxOutputStream);

				try (ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream()) {
					PdfOptions options = PdfOptions.create();
					PdfConverter.getInstance().convert(doc, pdfOutputStream, options);

					// gravar o arquivo pdf em disco pra testar mais facilmente
					OutputStream fileOutputStream = new FileOutputStream(pdfOutputPath);
					PdfConverter.getInstance().convert(doc, fileOutputStream, options);

					byte[] pdfBytes = Base64.getEncoder().encode(pdfOutputStream.toByteArray());
					String pdfBase64 = new String(pdfBytes);
					return new Resposta(pdfBase64, "teste");
				}
			}
		}
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
