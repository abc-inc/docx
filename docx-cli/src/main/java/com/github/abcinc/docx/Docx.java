package com.github.abcinc.docx;

import static java.nio.file.StandardOpenOption.READ;
import static org.slf4j.Logger.ROOT_LOGGER_NAME;
import static picocli.CommandLine.Command;
import static picocli.CommandLine.Option;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.Logger;
import ch.qos.logback.classic.LoggerContext;
import ch.qos.logback.classic.PatternLayout;
import ch.qos.logback.classic.spi.ILoggingEvent;
import ch.qos.logback.core.ConsoleAppender;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.slf4j.LoggerFactory;
import picocli.CommandLine;

@Command(
    name = "docx2",
    subcommands = {Convert.class, Slice.class},
    version = "docx 0.1.0",
    mixinStandardHelpOptions = true,
    description = "Process Word documents.")
public class Docx {

  static final org.slf4j.Logger logger = LoggerFactory.getLogger(Docx.class);
  static final ThreadLocal<CommandLine.ParseResult> parseResultThreadLocal =
      ThreadLocal.withInitial(() -> null);

  @Option(
      names = {"--font-dir"},
      description = "Additional font directory (can be specified multiple times).",
      hidden = true)
  @SuppressWarnings("unused")
  List<String> fontDirs = List.of();

  public static void main(String... args) {
    Docx docx = new Docx();
    System.exit(new CommandLine(docx).setExecutionStrategy(docx::executionStrategy).execute(args));
  }

  private int executionStrategy(CommandLine.ParseResult parseResult) {
    while (parseResult.hasSubcommand()) {
      parseResult = parseResult.subcommand();
    }

    LoggerContext lc = (LoggerContext) LoggerFactory.getILoggerFactory();
    PatternLayout layout = new PatternLayout();
    layout.setContext(lc);
    layout.setPattern("%msg%n");
    layout.start();

    ConsoleAppender<ILoggingEvent> appender =
        (ConsoleAppender<ILoggingEvent>) lc.getLogger(ROOT_LOGGER_NAME).getAppender("console");
    appender.setLayout(layout);

    String logLevel = parseResult.matchedOptionValue("--log-level", "");
    lc.getLogger(ROOT_LOGGER_NAME).setLevel(Level.toLevel(logLevel, Level.INFO));
    lc.getLogger("org.docx4j").setLevel(Level.toLevel(logLevel, Level.ERROR));

    /*
    XMLHelper always throws an error with the following stack trace:
    java.lang.IllegalArgumentException: Property 'http://javax.xml.XMLConstants/property/accessExternalSchema' is not recognized.
            at org.apache.xerces.jaxp.DocumentBuilderFactoryImpl.setAttribute(Unknown Source)
            at org.apache.poi.util.XMLHelper.trySet(XMLHelper.java:284)
            at org.apache.poi.util.XMLHelper.getDocumentBuilderFactory(XMLHelper.java:114)
            at org.apache.poi.util.XMLHelper.<clinit>(XMLHelper.java:85)
     */
    Logger logger = lc.getLogger("org.apache.poi.util.XMLHelper");
    logger.setLevel(Level.OFF);

    parseResultThreadLocal.set(parseResult);
    return new CommandLine.RunLast().execute(parseResult); // default execution strategy
  }

  static void outputWarning() {
    CommandLine.ParseResult parseResult = parseResultThreadLocal.get();
    Path output = parseResult.matchedOptionValue("--output", null);
    if (output != null) {
      return;
    }

    if (System.console() != null) {
      logger.warn(
          "Warning: Binary output can mess up your terminal. Use \"--output -\" to tell docx to output it to your terminal anyway, or consider \"--output <FILE>\" to save to a file.");
      System.exit(1);
    }
  }

  static OutputStream getOutputStream(Path outPath) throws IOException {
    if (outPath == null || outPath.toString().isEmpty()) {
      return System.out;
    }

    return Files.newOutputStream(outPath);
  }

  static InputStream getInputStream(Path inPath) throws IOException {
    if (inPath == null || inPath.toString().equals("-") || inPath.toString().equals("/dev/stdin")) {
      return System.in;
    }

    return Files.newInputStream(inPath, READ);
  }

  static void close(OutputStream out) {
    if (out == null || out == System.out) {
      return;
    }

    try {
      out.close();
    } catch (IOException e) {
      logger.warn("Failed to close output stream: {}", e.getMessage());
    }
  }
}
