package pocs.arquitetura.poc_wordtopdf;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

import lombok.SneakyThrows;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import javax.xml.bind.JAXBElement;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;

import java.util.Base64;
import java.util.List;

@RestController
public class PocController {

    private static final Logger logger = LoggerFactory.getLogger(PocController.class);

    @Autowired
    private DocxTemplateProvider docxTemplateProvider;

    @SneakyThrows
    @RequestMapping(
            value = "/pdf",
            method = RequestMethod.GET,
            produces = {
                    MediaType.APPLICATION_JSON_VALUE,
                    MediaType.APPLICATION_PDF_VALUE
            })
    public Object pdf(@RequestHeader("Accept") String acceptHeader) {

        // Com WordprocessingMLPackage é possível busar nodes com XPath,
        // com XWPFDocument ACHO que não (vale pesquisar pra não ter dois frameworks que fazer coisa similares).
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(docxTemplateProvider.getDocxTemplateFile());

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

        ByteArrayOutputStream docxOutputStream = new ByteArrayOutputStream();
        wordMLPackage.save(docxOutputStream);

        // transforma em objeto XWPFDocument para converter para PDF
        ByteArrayInputStream xwpfDocumentoInputStream = new ByteArrayInputStream(docxOutputStream.toByteArray());
        XWPFDocument document = new XWPFDocument(xwpfDocumentoInputStream);

        boolean debug = false;
        if (debug) { // salvar em disco para verificar o resultado

            String docxOutputPath = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato2.docx";
            String pdfOutputPath = "/home/zanfranceschi/projects/itau/poc_wordtopdf/output/contrato2.pdf";

            wordMLPackage.save(new File(docxOutputPath)); // salva o docx em disco

            // salva o pdf em disco
            try (FileOutputStream pdfOutputStream = new FileOutputStream(pdfOutputPath)) {
                PdfOptions options = PdfOptions.create();
                PdfConverter.getInstance().convert(document, pdfOutputStream, options);
            }
        }

        ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
        PdfOptions options = PdfOptions.create();
        PdfConverter.getInstance().convert(document, pdfOutputStream, options);

        byte[] pdfBytes = pdfOutputStream.toByteArray();

        pdfOutputStream.close();
        xwpfDocumentoInputStream.close();
        docxOutputStream.close();

        if (acceptHeader.equalsIgnoreCase(MediaType.APPLICATION_PDF_VALUE)) {

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_PDF);
            headers.setContentDispositionFormData("contrato-adesao", "contrato-adesao.pdf");
            headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
            ResponseEntity<byte[]> response = new ResponseEntity<>(pdfBytes, headers, HttpStatus.OK);
            return response;

        } else {

            byte[] pdfEncodedeBytes = Base64.getEncoder().encode(pdfBytes);
            String pdfBase64 = new String(pdfEncodedeBytes);
            return new Resposta(pdfBase64, "contrato-adesao");

        }
    }
}
