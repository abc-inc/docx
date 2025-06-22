package com.github.abcinc.docx;

import java.nio.file.Path;
import java.util.Collection;
import java.util.List;

public class Config {
  Path path;
  Path outPath;

  Collection<Integer> pages = List.of();

  public Config(Path path, Path outPath) {
    this.path = path;
    this.outPath = outPath;
  }

  public Path getPath() {
    return path;
  }

  public void setPath(Path path) {
    this.path = path;
  }

  public Path getOutPath() {
    return outPath;
  }

  public void setOutPath(Path outPath) {
    this.outPath = outPath;
  }

  public Collection<Integer> getPages() {
    return pages;
  }

  public void setPages(Collection<Integer> pages) {
    this.pages = pages;
  }
}
