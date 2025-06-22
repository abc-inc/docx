package com.github.abcinc.docx;

import static org.apache.commons.io.FilenameUtils.isExtension;

import jakarta.xml.bind.JAXBElement;
import jakarta.xml.bind.JAXBException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Stream;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.STBrType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Docx4JProcessor implements DocxProcessor {

  public static final Logger logger = LoggerFactory.getLogger(Docx4JProcessor.class);

  @Override
  public void convert(Config config, InputStream in, OutputStream out) throws IOException {
    List<String> dirs =
        List.of(
            "c:/windows/fonts",
            "c:/winnt/fonts",
            "d:/windows/fonts",
            "d:/winnt/fonts",
            "/usr/share/X11/fonts",
            "/usr/X/lib/X11/fonts",
            "/usr/openwin/lib/X11/fonts",
            "/usr/share/fonts",
            "/usr/X11R6/lib/X11/fonts",
            "/Library/Fonts",
            "/System/Library/Fonts",
            "/Applications/Microsoft Word.app/Contents/Resources/DFonts");

    dirs.stream().map(Paths::get).filter(Files::isDirectory).forEach(this::registerFonts);

    try {
      WordprocessingMLPackage p = Docx4J.load(in);
      Docx4J.pdfViaFO();
      Docx4J.toFO(new FOSettings(p), out, 0);
    } catch (Docx4JException e) {
      throw new IOException(e);
    }
  }

  @Override
  public void slice(Config config, InputStream in, OutputStream out) throws IOException {
    try {
      WordprocessingMLPackage p = Docx4J.load(in);
      MainDocumentPart mainDocumentPart = p.getMainDocumentPart();
      int currPage = 1;

      @SuppressWarnings("unchecked")
      List<P> paragraphs = (List) mainDocumentPart.getJAXBNodesViaXPath("//w:p", true);
      for (P paragraph : paragraphs) {
        if (!config.pages.isEmpty() && !config.pages.contains(currPage)) {
          mainDocumentPart.getJaxbElement().getContent().remove(paragraph);
        }

        for (Object content : List.copyOf(paragraph.getContent())) {
          R run = null;
          if (content instanceof R) {
            run = (R) content;
          } else if (content instanceof JAXBElement<?>) {
            Object conventValue = ((JAXBElement<?>) content).getValue();
            if (conventValue instanceof R) {
              run = (R) conventValue;
            }
          }
          if (run == null) {
            continue;
          }

          for (Object o : run.getContent()) {
            if (o instanceof Br && ((Br) o).getType() == STBrType.PAGE) {
              currPage++;
              logger.debug("Proceeding to page {}", currPage);
            }
          }
        }
      }

      while (mainDocumentPart.getJaxbElement().getContent().size() > 1) {
        Object last = mainDocumentPart.getJaxbElement().getContent().getLast();
        if (!(last instanceof P) || !isPageBreak((P) last)) {
          break;
        }
        mainDocumentPart.getJaxbElement().getContent().removeLast();
      }

      if (mainDocumentPart.getJaxbElement().getContent().isEmpty()) {
        throw new IOException("Document is empty after removing pages.");
      }

      p.save(out);
    } catch (Docx4JException | JAXBException e) {
      throw new IOException(e);
    }
  }

  boolean isPageBreak(P paragraph) {
    List<Object> content = paragraph.getContent();
    if (content == null || content.isEmpty() || !(content.getFirst() instanceof R)) {
      return false;
    }
    Object child = ((R) content.getFirst()).getContent().getFirst();
    return child instanceof Br && ((Br) child).getType() == STBrType.PAGE;
  }

  @Override
  public void registerFonts(Path dir) {
    try (Stream<Path> files = Files.list(dir)) {
      files
          .filter(Files::isRegularFile)
          .filter(f -> isExtension(f.getFileName().toString().toLowerCase(), "otf", "ttf", "ttc"))
          .map(Path::toUri)
          .sorted()
          .forEach(PhysicalFonts::addPhysicalFont);
    } catch (IOException e) {
      throw new UncheckedIOException("Failed to register fonts from directory: " + dir, e);
    }
  }
}
