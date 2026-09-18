import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:universal_html/html.dart' as html;

class InvoicesScreen extends StatefulWidget {
  const InvoicesScreen({super.key});

  @override
  State<InvoicesScreen> createState() => _InvoicesScreenState();
}

class _InvoicesScreenState extends State<InvoicesScreen> {
  Uint8List? _preparedFile;

  final TextEditingController _fileNameController =
      TextEditingController(text: 'فواتير');

  bool _isProcessing = false;
  String _status = 'ارفع ملف Excel للبدء';

  @override
  void dispose() {
    _fileNameController.dispose();
    super.dispose();
  }

  Future<void> _pickFile() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
      withData: true,
    );

    if (result == null || result.files.single.bytes == null) {
      return;
    }

    final bytes = result.files.single.bytes!;

    setState(() {
      _isProcessing = true;
      _preparedFile = null;
      _status = 'جاري معالجة الملف...';
    });

    try {
      final output = await _processFile(bytes);

      if (!mounted) return;

      setState(() {
        _preparedFile = output;
        _isProcessing = false;
        _status = 'تم تجهيز الملف بنجاح، اضغط تصدير الملف';
      });
    } catch (e) {
      if (!mounted) return;

      setState(() {
        _isProcessing = false;
        _preparedFile = null;
        _status = 'حدث خطأ أثناء معالجة الملف';
      });

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('خطأ: $e'),
          backgroundColor: Colors.red,
        ),
      );
    }
  }

  Future<Uint8List> _processFile(Uint8List inputBytes) async {
    final inputExcel = Excel.decodeBytes(inputBytes);

    Sheet? sourceSheet;

    if (inputExcel.tables.containsKey('فواتير')) {
      sourceSheet = inputExcel['فواتير'];
    } else if (inputExcel.tables.isNotEmpty) {
      sourceSheet = inputExcel[inputExcel.tables.keys.first];
    }

    if (sourceSheet == null || sourceSheet.maxRows < 2) {
      throw Exception('لم يتم العثور على بيانات الفواتير');
    }

    final rows = sourceSheet.rows;

    final headers = <String, int>{};

    for (int column = 0; column < rows.first.length; column++) {
      final value = _cellText(rows.first[column]).trim();
      if (value.isNotEmpty) {
        headers[_normalizeHeader(value)] = column;
      }
    }

    final dateIndex = _findHeader(headers, [
      'تاريخ الفاتورة',
      'التاريخ',
      'تاريخ',
    ]);

    final invoiceIndex = _findHeader(headers, [
      'رقم الفاتورة',
      'رقم فاتورة',
      'الفاتورة',
    ]);

    final solarQuantityIndex = _findHeader(headers, [
      'كمية سولار',
      'سولار',
    ]);

    final solarPriceIndex = _findHeader(headers, [
      'سعر سولار',
      'سعر السولار',
    ]);

    final quantity92Index = _findHeader(headers, [
      'كمية بنزين 92',
      'كمية بنزين92',
      'بنزين 92',
    ]);

    final price92Index = _findHeader(headers, [
      'سعر بنزين 92',
      'سعر بنزين92',
      'سعر 92',
    ]);

    final quantity95Index = _findHeader(headers, [
      'كمية بنزين 95',
      'كمية بنزين95',
      'بنزين 95',
    ]);

    final price95Index = _findHeader(headers, [
      'سعر بنزين 95',
      'سعر بنزين95',
      'سعر 95',
    ]);

    if (dateIndex == -1 ||
        invoiceIndex == -1 ||
        solarQuantityIndex == -1 ||
        solarPriceIndex == -1 ||
        quantity92Index == -1 ||
        price92Index == -1 ||
        quantity95Index == -1 ||
        price95Index == -1) {
      throw Exception(
        'لم يتم العثور على أعمدة الفواتير المطلوبة',
      );
    }

    final output = Excel.createExcel();

    final defaultSheet = output.getDefaultSheet();
    if (defaultSheet != null && defaultSheet != 'فواتير') {
      output.rename(defaultSheet, 'فواتير');
    }

    final sheet = output['فواتير'];
    sheet.isRTL = true;

    final titleStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Calibri),
      fontSize: 16,
      bold: true,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: ExcelColor.fromHexString('000000'),
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    );

    final headerStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Calibri),
      fontSize: 11,
      bold: true,
      fontColorHex: ExcelColor.black,
      backgroundColorHex: ExcelColor.fromHexString('D9D9D9'),
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      textWrapping: TextWrapping.WrapText,
      leftBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      rightBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      topBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      bottomBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
    );

    final dataStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Calibri),
      fontSize: 11,
      bold: true,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      rightBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      topBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      bottomBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
    );

    final totalStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Calibri),
      fontSize: 11,
      bold: true,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: ExcelColor.fromHexString('595959'),
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      rightBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      topBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      bottomBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
    );

    final summaryStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Calibri),
      fontSize: 11,
      bold: true,
      fontColorHex: ExcelColor.white,
      backgroundColorHex: ExcelColor.fromHexString('1F4E78'),
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      rightBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      topBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
      bottomBorder: Border(
        borderStyle: BorderStyle.Thin,
        borderColorHex: 'BFBFBF',
      ),
    );

    sheet.merge(
      CellIndex.indexByString('A1'),
      CellIndex.indexByString('L1'),
      customValue: TextCellValue(
        'اجمالي مسحوبات قسم نقل بورسعيد عن شهر اغسطس الفترة من 1/8/2026 الي 31/8/2026',
      ),
    );

    sheet.updateCell(
      CellIndex.indexByString('A1'),
      TextCellValue(
        'اجمالي مسحوبات قسم نقل بورسعيد عن شهر اغسطس الفترة من 1/8/2026 الي 31/8/2026',
      ),
      cellStyle: titleStyle,
    );

    sheet.setMergedCellStyle(
      CellIndex.indexByString('A1'),
      titleStyle,
    );

    sheet.setRowHeight(0, 30);

    const outputHeaders = [
      'تاريخ الفاتورة',
      'رقم الفاتورة',
      'كمية سولار',
      'سعر سولار',
      'قيمة سولار',
      'كمية بنزين 92',
      'سعر بنزين 92',
      'قيمة بنزين 92',
      'كمية بنزين 95',
      'سعر بنزين 95',
      'قيمة بنزين 95',
      'إجمالي القيمة',
    ];

    for (int column = 0; column < outputHeaders.length; column++) {
      sheet.updateCell(
        CellIndex.indexByColumnRow(column, 1),
        TextCellValue(outputHeaders[column]),
        cellStyle: headerStyle,
      );
    }

    sheet.setRowHeight(1, 24);

    final dataRows = <List<Object>>[];

    for (int rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      final row = rows[rowIndex];

      final dateValue = _cellValue(
        row,
        dateIndex,
      );

      final invoiceValue = _numberValue(
        row,
        invoiceIndex,
      );

      final solarQuantity = _numberValue(
        row,
        solarQuantityIndex,
      );

      final solarPrice = _numberValue(
        row,
        solarPriceIndex,
      );

      final quantity92 = _numberValue(
        row,
        quantity92Index,
      );

      final price92 = _numberValue(
        row,
        price92Index,
      );

      final quantity95 = _numberValue(
        row,
        quantity95Index,
      );

      final price95 = _numberValue(
        row,
        price95Index,
      );

      if (_isEmptyValue(dateValue) &&
          invoiceValue == 0 &&
          solarQuantity == 0 &&
          quantity92 == 0 &&
          quantity95 == 0) {
        continue;
      }

      dataRows.add([
        dateValue,
        invoiceValue,
        solarQuantity,
        solarPrice,
        solarQuantity * solarPrice,
        quantity92,
        price92,
        quantity92 * price92,
        quantity95,
        price95,
        quantity95 * price95,
        (solarQuantity * solarPrice) +
            (quantity92 * price92) +
            (quantity95 * price95),
      ]);
    }

    int excelRow = 3;

    for (final row in dataRows) {
      for (int column = 0; column < row.length; column++) {
        final value = row[column];

        sheet.updateCell(
          CellIndex.indexByColumnRow(column, excelRow - 1),
          _toCellValue(value),
          cellStyle: dataStyle,
        );
      }

      sheet.setRowHeight(excelRow - 1, 18);
      excelRow++;
    }

    final lastRow = excelRow - 1;
    final totalRow = lastRow + 1;

    sheet.updateCell(
      CellIndex.indexByColumnRow(0, totalRow - 1),
      TextCellValue('الإجمالي'),
      cellStyle: totalStyle,
    );

    const sumColumns = [
      2,
      4,
      5,
      7,
      8,
      10,
      11,
    ];

    for (int column = 0; column < 12; column++) {
      if (column != 0) {
        sheet.updateCell(
          CellIndex.indexByColumnRow(column, totalRow - 1),
          TextCellValue(''),
          cellStyle: totalStyle,
        );
      } else {
        sheet.updateCell(
          CellIndex.indexByColumnRow(column, totalRow - 1),
          TextCellValue('الإجمالي'),
          cellStyle: totalStyle,
        );
      }

      if (sumColumns.contains(column)) {
        final columnLetter = _columnLetter(column);

        sheet.updateCell(
          CellIndex.indexByColumnRow(column, totalRow - 1),
          FormulaCellValue(
            'SUM($columnLetter\$3:$columnLetter\$$lastRow)',
          ),
          cellStyle: totalStyle,
        );
      }
    }

    sheet.setRowHeight(totalRow - 1, 24);

    final summaryRow = totalRow + 1;

    final summaryValues = <int, Object>{
      1: 'سولار',
      2: FormulaCellValue('C\$$totalRow'),
      4: 'بنزين 92',
      5: FormulaCellValue('F\$$totalRow'),
      7: 'بنزين 95',
      8: FormulaCellValue('I\$$totalRow'),
    };

    for (int column = 0; column < 12; column++) {
      final value = summaryValues[column] ?? '';

      sheet.updateCell(
        CellIndex.indexByColumnRow(column, summaryRow - 1),
        _toCellValue(value),
        cellStyle: summaryStyle,
      );
    }

    sheet.setRowHeight(summaryRow - 1, 22);

    const widths = [
      16.0,
      14.0,
      14.0,
      12.0,
      15.0,
      16.0,
      16.0,
      18.0,
      16.0,
      16.0,
      18.0,
      18.0,
    ];

    for (int column = 0; column < widths.length; column++) {
      sheet.setColumnWidth(column, widths[column]);
    }

    final encoded = output.encode();

    if (encoded == null || encoded.isEmpty) {
      throw Exception('تعذر إنشاء ملف Excel');
    }

    return Uint8List.fromList(encoded);
  }

  Future<void> _exportFile() async {
    if (_preparedFile == null) return;

    String fileName = _fileNameController.text.trim();

    if (fileName.isEmpty) {
      fileName = 'فواتير';
    }

    if (!fileName.toLowerCase().endsWith('.xlsx')) {
      fileName = '$fileName.xlsx';
    }

    final blob = html.Blob(
      [_preparedFile!],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      ..setAttribute('download', fileName)
      ..style.display = 'none';

    html.document.body?.append(anchor);
    anchor.click();
    anchor.remove();

    html.Url.revokeObjectUrl(url);
  }

  int _findHeader(
    Map<String, int> headers,
    List<String> names,
  ) {
    for (final name in names) {
      final index = headers[_normalizeHeader(name)];
      if (index != null) {
        return index;
      }
    }

    return -1;
  }

  String _normalizeHeader(String value) {
    return value
        .replaceAll('\u00A0', ' ')
        .replaceAll(RegExp(r'\s+'), ' ')
        .trim()
        .toLowerCase();
  }

  String _cellText(Data? cell) {
    if (cell == null || cell.value == null) {
      return '';
    }

    return cell.value.toString();
  }

  Object? _cellValue(List<Data?> row, int index) {
    if (index < 0 || index >= row.length) {
      return null;
    }

    final value = row[index]?.value;

    if (value is TextCellValue) {
      return value.value.text;
    }

    return value;
  }

  double _numberValue(List<Data?> row, int index) {
    final value = _cellValue(row, index);

    if (value == null) {
      return 0;
    }

    if (value is IntCellValue) {
      return value.value.toDouble();
    }

    if (value is DoubleCellValue) {
      return value.value;
    }

    if (value is FormulaCellValue) {
      final result = value.result;

      if (result is num) {
        return result.toDouble();
      }

      return double.tryParse(
            result?.toString().replaceAll(',', '').trim() ?? '',
          ) ??
          0;
    }

    if (value is num) {
      return value.toDouble();
    }

    final text = value
        .toString()
        .replaceAll(',', '')
        .replaceAll('٬', '')
        .replaceAll('،', '.')
        .trim();

    return double.tryParse(text) ?? 0;
  }

  bool _isEmptyValue(Object? value) {
    if (value == null) return true;
    if (value.toString().trim().isEmpty) return true;
    return false;
  }

  CellValue _toCellValue(Object? value) {
    if (value is CellValue) {
      return value;
    }

    if (value is int) {
      return IntCellValue(value);
    }

    if (value is double) {
      return DoubleCellValue(value);
    }

    if (value is num) {
      return DoubleCellValue(value.toDouble());
    }

    return TextCellValue(value?.toString() ?? '');
  }

  String _columnLetter(int index) {
    int number = index + 1;
    String result = '';

    while (number > 0) {
      final remainder = (number - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      number = (number - 1) ~/ 26;
    }

    return result;
  }

  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: TextDirection.rtl,
      child: Scaffold(
        appBar: AppBar(
          title: const Text(
            'الفواتير',
            style: TextStyle(
              fontWeight: FontWeight.bold,
            ),
          ),
          centerTitle: true,
          backgroundColor: Colors.blue,
          foregroundColor: Colors.white,
        ),
        body: Center(
          child: ConstrainedBox(
            constraints: const BoxConstraints(
              maxWidth: 600,
            ),
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(24),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  const Icon(
                    Icons.receipt_long_rounded,
                    size: 64,
                    color: Colors.blue,
                  ),
                  const SizedBox(height: 18),
                  const Text(
                    'كشف الفواتير',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 25,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                  const SizedBox(height: 8),
                  Text(
                    'اختر ملف Excel وسيتم تجهيزه تلقائيًا',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 14,
                      color: Colors.grey.shade600,
                    ),
                  ),
                  const SizedBox(height: 28),
                  OutlinedButton.icon(
                    onPressed: _isProcessing ? null : _pickFile,
                    icon: _isProcessing
                        ? const SizedBox(
                            width: 20,
                            height: 20,
                            child: CircularProgressIndicator(
                              strokeWidth: 2,
                            ),
                          )
                        : const Icon(Icons.upload_file_rounded),
                    label: Text(
                      _isProcessing
                          ? 'جاري المعالجة...'
                          : 'اختيار ملف Excel',
                    ),
                    style: OutlinedButton.styleFrom(
                      padding: const EdgeInsets.symmetric(
                        vertical: 16,
                      ),
                    ),
                  ),
                  const SizedBox(height: 18),
                  Container(
                    padding: const EdgeInsets.all(16),
                    decoration: BoxDecoration(
                      color: _preparedFile != null
                          ? Colors.green.withValues(alpha: 0.08)
                          : Colors.grey.withValues(alpha: 0.08),
                      borderRadius: BorderRadius.circular(12),
                      border: Border.all(
                        color: _preparedFile != null
                            ? Colors.green.withValues(alpha: 0.25)
                            : Colors.grey.withValues(alpha: 0.20),
                      ),
                    ),
                    child: Text(
                      _status,
                      textAlign: TextAlign.center,
                      style: TextStyle(
                        fontWeight: FontWeight.w600,
                        color: _preparedFile != null
                            ? Colors.green.shade700
                            : Colors.grey.shade700,
                      ),
                    ),
                  ),
                  const SizedBox(height: 18),
                  TextField(
                    controller: _fileNameController,
                    textDirection: TextDirection.rtl,
                    decoration: InputDecoration(
                      labelText: 'اسم الملف',
                      hintText: 'فواتير',
                      prefixIcon: const Icon(Icons.edit_document),
                      border: OutlineInputBorder(
                        borderRadius: BorderRadius.circular(12),
                      ),
                    ),
                  ),
                  const SizedBox(height: 18),
                  FilledButton.icon(
                    onPressed: _preparedFile == null
                        ? null
                        : _exportFile,
                    icon: const Icon(Icons.download_rounded),
                    label: const Text('تصدير الملف'),
                    style: FilledButton.styleFrom(
                      padding: const EdgeInsets.symmetric(
                        vertical: 16,
                      ),
                    ),
                  ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}
```0
