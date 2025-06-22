package com.github.abcinc.docx;

import static com.github.abcinc.docx.Docx.close;
import static com.github.abcinc.docx.Docx.getInputStream;
import static com.github.abcinc.docx.Docx.getOutputStream;
import static com.github.abcinc.docx.Docx.outputWarning;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.concurrent.Callable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import picocli.CommandLine;

@CommandLine.Command(
    name = "convert",
    mixinStandardHelpOptions = true,
    description = "Convert Word documents to a different format.")
public class Convert implements Callable<Integer> {

  static final Logger logger = LoggerFactory.getLogger(Convert.class);

  @CommandLine.Option(
      names = {"--log-level"},
      description = "Set the logging level (default: WARN).",
      defaultValue = "WARN")
  @SuppressWarnings("unused")
  String logLevel;

  @CommandLine.Option(
      names = {"-f", "--format"},
      description = "Output format.",
      hidden = true)
      @SuppressWarnings("unused")
  String format = "pdf";

  @CommandLine.Option(
      names = {"-o", "--output"},
      type = Path.class,
      description = "Output PDF file.")
  Path outPath;

  @CommandLine.Option(
      names = {"--processor"},
      description = "Processor to use for conversion.",
      defaultValue = "docx4j")
  String processor;

  @CommandLine.Parameters(index = "0", description = "Word document to load.", defaultValue = "-")
  Path path;

  @Override
  public Integer call() throws IOException {
    outputWarning();

    Config config = new Config(path, outPath);
    OutputStream out = getOutputStream(outPath);

    try (InputStream in = getInputStream(path)) {
      switch (processor) {
        case "docx4j" -> new Docx4JProcessor().convert(config, in, out);
        case "poi" -> new POIProcessor().convert(config, in, out);
        default -> throw new IllegalArgumentException("Unsupported processor: " + processor);
      }
      return 0;
    } catch (Exception e) {
      logger.error(e.getMessage(), e);
    } finally {
      close(out);
    }
    return 1;
  }
}
