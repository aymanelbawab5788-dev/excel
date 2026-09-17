import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:universal_html/html.dart' as html;

class Calculation530Screen extends StatefulWidget {
  const Calculation530Screen({super.key});

  @override
  State<Calculation530Screen> createState() => _Calculation530ScreenState();
}

class _Calculation530ScreenState extends State<Calculation530Screen> {
  bool _processing = false;
  String _status = '';

  Future<void> _processFile() async {
    setState(() {
      _processing = true;
      _status = 'جاري قراءة الملف...';
    });

    try {
      final file = await FilePicker.pickFile(
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (file == null) {
        setState(() {
          _processing = false;
          _status = 'تم إلغاء اختيار الملف.';
        });
        return;
      }

      final bytes = await file.readAsBytes();

      setState(() {
        _status = 'جاري معالجة البيانات...';
      });

      final outputBytes = _buildExcel(bytes);

      _downloadFile(
        outputBytes,
        'تفريغ_مسحوبات_السولار_شهر_8.xlsx',
      );

      setState(() {
        _processing = false;
        _status = 'تم تجهيز الملف وتنزيله بنجاح.';
      });
    } catch (e) {
      setState(() {
        _processing = false;
        _status = 'حدث خطأ: $e';
      });
    }
  }

  Uint8List _buildExcel(Uint8List inputBytes) {
    final inputExcel = Excel.decodeBytes(inputBytes);

    if (!inputExcel.tables.containsKey('Report')) {
      throw Exception('الملف لا يحتوي على شيت باسم Report.');
    }

    final reportSheet = inputExcel.tables['Report']!;

    final List<Map<String, dynamic>> rawRows = [];

    for (int r = 1; r < reportSheet.maxRows; r++) {
      final carNumValue = _cellValue(reportSheet, r, 0);

      if (carNumValue == null || carNumValue.toString().trim().isEmpty) {
        continue;
      }

      final carNum = _formatCarNumber(carNumValue);

      rawRows.add({
        'car_num': carNum,
        'car_letters': _cellString(reportSheet, r, 1),
        'station': _cellString(reportSheet, r, 2).trim(),
        'date': _cellString(reportSheet, r, 3),
        'counter': _cellNumber(reportSheet, r, 4),
        'liters': _cellNumber(reportSheet, r, 5),
        'pct': _cellNumber(reportSheet, r, 6),
      });
    }

    final List<List<Map<String, dynamic>>> carGroups = [];

    List<Map<String, dynamic>> currentGroup = [];
    String? lastKey;

    for (final record in rawRows) {
      final key =
          '${record['car_num']}|${record['car_letters']}';

      if (key != lastKey) {
        if (currentGroup.isNotEmpty) {
          carGroups.add(currentGroup);
        }

        currentGroup = [record];
        lastKey = key;
      } else {
        currentGroup.add(record);
      }
    }

    if (currentGroup.isNotEmpty) {
      carGroups.add(currentGroup);
    }

    final outputExcel = Excel.createExcel();

    final defaultSheetName = outputExcel.getDefaultSheet();

    if (defaultSheetName != null &&
        outputExcel.tables.containsKey(defaultSheetName)) {
      outputExcel.delete(defaultSheetName);
    }

    final sheet = outputExcel['سولار شهر 8'];

    const headers = [
      'م',
      'رقم السيارة',
      'الأحرف',
      'المحطة',
      'التاريخ',
      'العداد',
      'اللترات',
      'النسبة الفعلية',
    ];

    for (int col = 0; col < headers.length; col++) {
      sheet
          .cell(CellIndex.indexByColumnRow(
            columnIndex: col,
            rowIndex: 1,
          ))
          .value = TextCellValue(headers[col]);
    }

    int currentRow = 2;

    for (int groupIndex = 0;
        groupIndex < carGroups.length;
        groupIndex++) {
      final group = carGroups[groupIndex];

      final startRow = currentRow;

      for (final record in group) {
        final values = [
          '',
          record['car_num'],
          record['car_letters'],
          record['station'],
          record['date'],
          record['counter'],
          record['liters'],
          record['pct'],
        ];

        for (int col = 0; col < values.length; col++) {
          final cell = sheet.cell(
            CellIndex.indexByColumnRow(
              columnIndex: col,
              rowIndex: currentRow,
            ),
          );

          final value = values[col];

          if (value is num) {
            cell.value = DoubleCellValue(value.toDouble());
          } else {
            cell.value = TextCellValue(value?.toString() ?? '');
          }
        }

        currentRow++;
      }

      final endDataRow = currentRow - 1;

      final serialCell = sheet.cell(
        CellIndex.indexByColumnRow(
          columnIndex: 0,
          rowIndex: startRow,
        ),
      );

      serialCell.value =
          IntCellValue(groupIndex + 1);

      final sumCell = sheet.cell(
        CellIndex.indexByColumnRow(
          columnIndex: 6,
          rowIndex: currentRow,
        ),
      );

      sumCell.value = FormulaCellValue(
        'SUM(G${startRow + 1}:G${endDataRow + 1})',
      );

      currentRow++;
    }

    final encoded = outputExcel.encode();

    if (encoded == null) {
      throw Exception('فشل إنشاء ملف Excel.');
    }

    return Uint8List.fromList(encoded);
  }

  dynamic _cellValue(
    Sheet sheet,
    int row,
    int column,
  ) {
    final cell = sheet.cell(
      CellIndex.indexByColumnRow(
        columnIndex: column,
        rowIndex: row,
      ),
    );

    final value = cell.value;

    if (value == null) {
      return null;
    }

    return value;
  }

  String _cellString(
    Sheet sheet,
    int row,
    int column,
  ) {
    final value = _cellValue(sheet, row, column);

    if (value == null) {
      return '';
    }

    return value.toString();
  }

  double _cellNumber(
    Sheet sheet,
    int row,
    int column,
  ) {
    final value = _cellValue(sheet, row, column);

    if (value == null) {
      return 0;
    }

    if (value is IntCellValue) {
      return value.value.toDouble();
    }

    if (value is DoubleCellValue) {
      return value.value;
    }

    final parsed = double.tryParse(value.toString());

    return parsed ?? 0;
  }

  String _formatCarNumber(dynamic value) {
    if (value is IntCellValue) {
      return value.value.toString();
    }

    if (value is DoubleCellValue) {
      final number = value.value;

      if (number == number.truncateToDouble()) {
        return number.toInt().toString();
      }

      return number.toString();
    }

    final text = value.toString();

    final number = double.tryParse(text);

    if (number != null &&
        number == number.truncateToDouble()) {
      return number.toInt().toString();
    }

    return text;
  }

  void _downloadFile(
    Uint8List bytes,
    String fileName,
  ) {
    final blob = html.Blob(
      [bytes],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    final url = html.Url.createObjectUrlFromBlob(blob);

    final anchor = html.AnchorElement(href: url)
      ..setAttribute('download', fileName)
      ..click();

    html.Url.revokeObjectUrl(url);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('حساب 530'),
      ),
      body: Center(
        child: Padding(
          padding: const EdgeInsets.all(24),
          child: ConstrainedBox(
            constraints: const BoxConstraints(
              maxWidth: 500,
            ),
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                const Icon(
                  Icons.table_view,
                  size: 64,
                ),
                const SizedBox(height: 20),
                const Text(
                  'حساب 530',
                  style: TextStyle(
                    fontSize: 24,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                const SizedBox(height: 12),
                const Text(
                  'اختر ملف Excel الذي يحتوي على شيت Report',
                  textAlign: TextAlign.center,
                ),
                const SizedBox(height: 24),
                SizedBox(
                  width: double.infinity,
                  child: ElevatedButton.icon(
                    onPressed:
                        _processing ? null : _processFile,
                    icon: const Icon(Icons.upload_file),
                    label: Text(
                      _processing
                          ? 'جاري المعالجة...'
                          : 'اختيار ملف Excel',
                    ),
                  ),
                ),
                if (_status.isNotEmpty) ...[
                  const SizedBox(height: 20),
                  Text(
                    _status,
                    textAlign: TextAlign.center,
                  ),
                ],
              ],
            ),
          ),
        ),
      ),
    );
  }
}
