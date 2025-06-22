# docx

*docx* is a Java command line tool for processing Microsoft Word documents in the DOCX format.
It provides functionality to slice documents as well as convert documents to PDF format.
Different processors like docx4j and Apache POI are supported for these operations.
Hence, it does not require any additional dependencies to be installed on the system.

## Usage

```shell
java -jar docx-cli-all.jar slice --pages 1-2 --output output.docx input.docx
java -jar docx-cli-all.jar convert --output output.pdf input.docx
```

## Native Build

*docx* is provided as a self-contained jar file,
but it can also be built as a native executable using [GraalVM](https://www.graalvm.org/).

To build the native executable, you need to have GraalVM with AWT support installed on your system.
Therefore, it is recommended to install [BellSoft Liberica JDK](https://bell-sw.com/libericajdk/).

It can also be installed using [SDKMAN!](https://sdkman.io/):

```shell
sdk install java 23.1.7.r21-nik
```

Then, you can build the native executable by running the following commands:

```shell
# ensure JAVA_HOME is set to the correct GraalVM installation
./gradlew build
./gradlew -Pagent run --args 'convert ../testdata/fonts.docx -o build/tmp/docx4j.pdf --processor docx4j'
./gradlew metadataCopy
./gradlew -Pagent run --args 'convert ../testdata/fonts.docx -o build/tmp/poi.pdf --processor poi'
./gradlew metadataCopy
./gradlew -Pagent run --args 'slice ../testdata/fonts.docx -o build/tmp/docx4j.docx --processor docx4j'
./gradlew metadataCopy
./gradlew -Pagent run --args 'slice ../testdata/fonts.docx -o build/tmp/poi.docx --processor poi'
./gradlew metadataCopy
./gradlew nativeCompile
```
