package com.github.abcinc.docx;

import static com.github.abcinc.docx.Docx.getInputStream;
import static com.github.abcinc.docx.Docx.outputWarning;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.Callable;
import picocli.CommandLine;

@CommandLine.Command(
    name = "slice",
    mixinStandardHelpOptions = true,
    description = "Returns the pages in the range specified.")
public class Slice implements Callable<Integer> {

  @CommandLine.Option(
      names = {"--log-level"},
      description = "Set the logging level (default: WARN).",
      defaultValue = "WARN")
  @SuppressWarnings("unused")
  String logLevel;

  @CommandLine.Option(
      names = {"-o", "--output"},
      type = Path.class,
      description = "Output PDF file.")
  Path outPath;

  @CommandLine.Option(
      names = {"-p", "--pages"},
      description = "Pages to extract.")
  List<String> pages = List.of();

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
    config.setPages(Slicing.rangesToList(Slicing.calcPages(pages)));

    OutputStream out = Docx.getOutputStream(outPath);
    try (InputStream in = getInputStream(path)) {
      switch (processor) {
        case "docx4j" -> new Docx4JProcessor().slice(config, in, out);
        case "poi" -> new POIProcessor().slice(config, in, out);
        default -> throw new IllegalArgumentException("Unsupported processor: " + processor);
      }
      return 0;
    }
  }
}
