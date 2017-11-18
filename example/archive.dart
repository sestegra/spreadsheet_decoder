import 'dart:io';
import 'dart:typed_data';
import 'package:archive/archive.dart';

Archive cloneArchive(Archive archive) {
  var clone = new Archive();
  archive.files.forEach((file) {
    if (file.isFile) {
      var content = (file.content as Uint8List).toList();
      var copy = new ArchiveFile(file.name, content.length, content)..compress = file.compress;
      clone.addFile(copy);
    }
  });
  return clone;
}

void unzip(String path) {
  var unzip = Process.runSync("unzip", ["-t", path]);
  print(unzip.stdout);
  print(unzip.stderr);
}

void binwalk(String path) {
  var binwalk = Process.runSync("binwalk", [path]);
  print(binwalk.stdout);
  print(binwalk.stderr);
}

void main(List<String> args) {
  var path = args[0];

  for (var file in new Directory(path).listSync(recursive: false)) {
    if (file is File) {
      var input = file.path;
      print(input);
      var archive = new ZipDecoder().decodeBytes(new File(input).readAsBytesSync(), verify: true);
      var zip = new ZipEncoder().encode(cloneArchive(archive));
      try {
        archive = new ZipDecoder().decodeBytes(zip, verify: true);
        var file = new File("out/$input")
          ..createSync(recursive: true)
          ..writeAsBytesSync(zip);
        unzip(file.path);
        binwalk(file.path);

      } catch (e) {
        print("$input: KO");
      }
    }
  }
}
