import 'dart:typed_data';
import 'package:excel/excel.dart';

class ExcelFileInfo {
  ExcelFileInfo({required this.name, required this.bytes});
  final String name;
  final Uint8List bytes;
  String get sizeLabel => bytes.lengthInBytes < 1024 ? '${bytes.lengthInBytes} بايت' : '${(bytes.lengthInBytes / 1024).toStringAsFixed(1)} كيلوبايت';
}

class ProcessingReport {
  ProcessingReport({required this.bytes, required this.fileName, required this.summary});
  final Uint8List bytes; final String fileName; final String summary;
}

class ExcelProcessor {
  ProcessingReport process({required Uint8List referenceBytes, required Uint8List dataBytes, required String sourceName, void Function(String)? onStatus}) {
    onStatus?.call('جاري تحليل الملف المرجعي');
    final reference = Excel.decodeBytes(referenceBytes);
    onStatus?.call('جاري تحليل ملف البيانات الجديدة');
    final incoming = Excel.decodeBytes(dataBytes);
    final referenceShape = _inspect(_rows(_bestSheet(reference, incoming)));
    final dataShape = _inspect(_rows(_bestSheet(incoming, reference)));
    if (referenceShape.headers.isEmpty || dataShape.headers.isEmpty) throw const FormatException('تعذر اكتشاف صف عناوين واضح في أحد الملفين.');
    onStatus?.call('جاري تطبيق النموذج');
    final rows = _buildRows(referenceShape, dataShape);
    final output = Excel.createExcel();
    final sheet = output[output.getDefaultSheet() ?? 'Sheet1'];
    for (var row = 0; row < rows.length; row++) {
      for (var column = 0; column < rows[row].length; column++) {
        final value = rows[row][column];
        if (value != null) sheet.cell(CellIndex.indexByColumnRow(columnIndex: column, rowIndex: row)).value = value;
      }
    }
    for (var column = 0; column < referenceShape.headers.length; column++) {
      sheet.setColumnWidth(column, column == 0 ? 20 : 16);
    }
    onStatus?.call('جاري إنشاء الملف النهائي');
    final encoded = output.encode();
    if (encoded == null) throw const FormatException('تعذر إنشاء ملف Excel النهائي.');
    final grouped = referenceShape.groupColumn != null && referenceShape.quantityColumn != null;
    return ProcessingReport(bytes: Uint8List.fromList(encoded), fileName: '${_baseName(sourceName)}_منسق.xlsx', summary: grouped ? 'تم تنسيق البيانات وتجميع الكميات حسب المجموعة.' : 'تم تنسيق البيانات وتطبيق ترتيب المرجع.');
  }

  List<List<dynamic>> _rows(Sheet sheet) => sheet.rows.map((row) => row.map((cell) => cell?.value).toList()).toList();
  Sheet _bestSheet(Excel first, Excel second) {
    final comparison = _inspect(_rows(second.tables.values.first)); var best = first.tables.values.first; var score = -1;
    for (final sheet in first.tables.values) { final current = _inspect(_rows(sheet)); final headers = comparison.headers.map(_key).toSet(); final candidate = current.headers.map(_key).where(headers.contains).length * 10 - (current.headers.length - comparison.headers.length).abs(); if (candidate > score) { best = sheet; score = candidate; } }
    return best;
  }
  _Shape _inspect(List<List<dynamic>> rows) {
    if (rows.isEmpty) return const _Shape([], [], null, null); var headerIndex = 0; var score = -1;
    for (var index = 0; index < rows.length && index < 30; index++) { final nonEmpty = rows[index].where((value) => value != null && value.toString().trim().isNotEmpty).length; final text = rows[index].where((value) => value is String && value.toString().trim().isNotEmpty).length; if (nonEmpty + text * 2 > score) { score = nonEmpty + text * 2; headerIndex = index; } }
    final headers = rows[headerIndex].map((value) => value?.toString().trim() ?? '').toList(); final data = rows.skip(headerIndex + 1).where((row) => row.any((value) => value != null && value.toString().trim().isNotEmpty)).toList();
    final group = _findColumn(headers, const ['سيارة', 'مركبة', 'vehicle', 'car', 'معرف', 'رقم']); final quantity = _findColumn(headers, const ['كمية', 'لتر', 'liter', 'litre', 'quantity', 'سحب', 'amount']); final repeated = group != null && data.map((row) => _value(row, group)).where((value) => value != null).toSet().length < data.length;
    return _Shape(headers, data, repeated ? group : null, quantity);
  }
  List<List<dynamic>> _buildRows(_Shape reference, _Shape incoming) {
    final mapped = incoming.data.map((row) => reference.headers.map((header) { final index = incoming.headers.indexWhere((candidate) => _key(candidate) == _key(header)); return index < 0 ? null : _value(row, index); }).toList()).toList(); final result = <List<dynamic>>[reference.headers];
    if (reference.groupColumn == null || reference.quantityColumn == null) return result..addAll(mapped); final grouped = <String, List<List<dynamic>>>{};
    for (final row in mapped) { final key = _value(row, reference.groupColumn!); if (key != null) grouped.putIfAbsent(key.toString(), () => []).add(row); }
    for (final items in grouped.values) { final merged = List<dynamic>.from(items.first); merged[reference.quantityColumn!] = items.map((row) => _number(_value(row, reference.quantityColumn!)) ?? 0).fold<num>(0, (sum, value) => sum + value); result.add(merged); }
    return result;
  }
  int? _findColumn(List<String> headers, List<String> words) {
    for (var i = 0; i < headers.length; i++) {
      if (words.any((word) => _key(headers[i]).contains(_key(word)))) return i;
    }
    return null;
  }
  dynamic _value(List<dynamic> row, int index) => index < row.length ? row[index] : null;
  num? _number(dynamic value) => value is num ? value : num.tryParse(value?.toString().replaceAll(',', '').trim() ?? '');
  String _key(String value) => value.toLowerCase().replaceAll(RegExp(r'[\s_\-/:]+'), '').replaceAll('أ', 'ا').replaceAll('إ', 'ا').replaceAll('آ', 'ا');
  String _baseName(String value) => value.replaceFirst(RegExp(r'\.[^.]+$'), '');
}

class _Shape {
  const _Shape(this.headers, this.data, this.groupColumn, this.quantityColumn);
  final List<String> headers; final List<List<dynamic>> data; final int? groupColumn; final int? quantityColumn;
}