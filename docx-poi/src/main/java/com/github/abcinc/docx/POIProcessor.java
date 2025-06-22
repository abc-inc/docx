package com.github.abcinc.docx;

import static com.google.common.collect.Iterables.getOnlyElement;

import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBrType;

public class POIProcessor implements DocxProcessor {

  @Override
  public void convert(Config config, InputStream in, OutputStream out) throws IOException {
    FontFactory.registerDirectories();
    FontFactory.registerDirectory("/Applications/Microsoft Word.app/Contents/Resources/DFonts");

    try (XWPFDocument doc = new XWPFDocument(in)) {
      PdfOptions opts = PdfOptions.getDefault().fontProvider(new CustomFontProvider());
      PdfConverter.getInstance().convert(doc, out, opts);
    } catch (XWPFConverterException e) {
      throw new IOException(e);
    }
  }

  @Override
  public void slice(Config config, InputStream in, OutputStream out) throws IOException {
    try (XWPFDocument doc = new XWPFDocument(in)) {
      int currPage = 1;
      List<XWPFParagraph> paragraphs = doc.getParagraphs();
      List<Integer> removed = new ArrayList<>();
      for (int i = 0; i < paragraphs.size(); i++) {
        XWPFParagraph paragraph = paragraphs.get(i);
        if (!config.pages.isEmpty() && !config.pages.contains(currPage)) {
          removed.add(i);
        }

        for (XWPFRun run : paragraph.getRuns()) {
          List<CTBr> brList = run.getCTR().getBrList();
          if (!brList.isEmpty() && getOnlyElement(brList).getType() == STBrType.PAGE) {
            currPage++;
          }
        }
      }

      for (int i = removed.size() - 1; i >= 0; i--) {
        doc.removeBodyElement(removed.get(i));
      }

      doc.write(out);
    } catch (XWPFConverterException e) {
      throw new IOException(e);
    }
  }

  @Override
  public void registerFonts(Path path) {
    FontFactory.registerDirectory(path.toString());
  }

  private static class CustomFontProvider implements IFontProvider {
    @Override
    public Font getFont(String familyName, String encoding, float size, int style, Color color) {
      // Use CP1252 encoding, because IDENTITY_H does not work as expected.
      return FontFactory.getFont(familyName, BaseFont.CP1252, size, style, color);
    }
  }
}
