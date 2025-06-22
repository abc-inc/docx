package com.github.abcinc.docx;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;

public interface DocxProcessor {

  void convert(Config config, InputStream in, OutputStream out) throws IOException;

  void slice(Config config, InputStream in, OutputStream out) throws IOException;

  void registerFonts(Path dir);
}
