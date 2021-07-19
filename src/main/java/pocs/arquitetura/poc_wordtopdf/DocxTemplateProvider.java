package pocs.arquitetura.poc_wordtopdf;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileNotFoundException;
import java.time.LocalDateTime;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;

@Service
public class DocxTemplateProvider {

    private static final Logger logger = LoggerFactory.getLogger(DocxTemplateProvider.class);

    private static final String templatePath = DocxTemplateProvider.class.getResource("/contrato-template.docx").getPath();

    private File docxTemplateFile;
    private LocalDateTime templateLoadedAt;
    private int templateExpirationSeconds = 60; // 2 * 60 * 60; // 2 horas

    public File getDocxTemplateFile() {

        // recarrega o template de tempos em tempos (definido em 'templateExpirationSeconds')
        if (
                docxTemplateFile == null ||
                templateLoadedAt == null ||
                templateLoadedAt.plusSeconds(templateExpirationSeconds).isBefore(LocalDateTime.now())
        ) {
            docxTemplateFile = new File(templatePath);
            templateLoadedAt = LocalDateTime.now();
            logger.debug("|===============================|");
            logger.debug("| template carregado/atualizado |");
            logger.debug("|===============================|");
        }

        return docxTemplateFile;
    }
}
