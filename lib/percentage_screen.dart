import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:universal_html/html.dart' as html;

class PercentageScreen extends StatefulWidget {
  const PercentageScreen({super.key});

  @override
  State<PercentageScreen> createState() => _PercentageScreenState();
}

class _PercentageScreenState extends State<PercentageScreen> {
  Uint8List? _selectedBytes;
  String? _selectedFileName;

  bool _isProcessing = false;
  bool _isReady = false;

  String _status = 'ارفع ملف Excel للبدء';

  Future<void> _pickFile() async {
    try {
      final file = await FilePicker.pickFile(
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
      );

      if (file == null) {
        return;
      }

      final bytes = await file.readAsBytes();

      setState(() {
        _selectedBytes = bytes;
        _selectedFileName = file.name;
        _isReady = false;
        _status = 'تم اختيار الملف، جاهز للمعالجة';
      });
    } catch (e) {
      _showError('حدث خطأ أثناء اختيار الملف:\n$e');
    }
  }

  Future<void> _processFile() async {
    if (_selectedBytes == null) {
      _showError('اختر ملف Excel أولاً');
      return;
    }

    setState(() {
      _isProcessing = true;
      _isReady = false;
      _status = 'جاري تجهيز كشف النسبة...';
    });

    try {
      final input = Excel.decodeBytes(_selectedBytes!);

      if (!input.tables.containsKey('Report')) {
        throw const FormatException(
          'الملف لا يحتوي على Sheet باسم Report',
        );
      }

      final reportSheet = input.tables['Report']!;

      final detailsSheet = input.tables['تفاصيل التموينات'];

      final detailsMap = _buildDetailsMap(detailsSheet);

      final output = Excel.createExcel();
      final outputSheet = output['Sheet1'];

      _buildOutput(
        outputSheet,
        reportSheet,
        detailsMap,
      );

      setState(() {
        _isReady = true;
        _status = 'تم تجهيز كشف النسبة بنجاح';
      });

      final encoded = output.encode();

      if (encoded == null) {
        throw const FormatException(
          'تعذر إنشاء ملف Excel النهائي',
        );
      }

      _downloadFile(
        Uint8List.fromList(encoded),
        'كشف نسبة.xlsx',
      );
    } catch (e) {
      setState(() {
        _status = 'حدث خطأ أثناء المعالجة';
      });

      _showError('تعذر معالجة الملف:\n$e');
    } finally {
      if (mounted) {
        setState(() {
          _isProcessing = false;
        });
      }
    }
  }

  Map<String, Map<String, dynamic>> _buildDetailsMap(
    Sheet? sheet,
  ) {
    final result = <String, Map<String, dynamic>>{};

    if (sheet == null) {
      return result;
    }

    final rows = sheet.rows;

    if (rows.isEmpty) {
      return result;
    }

    final headers = _findHeaders(rows);

    if (headers.isEmpty) {
      return result;
    }

    final vehicleIndex = headers['رقم السيارة'];

    if (vehicleIndex == null) {
      return result;
    }

    for (var i = 1; i < rows.length; i++) {
      final row = rows[i];

      final vehicle = _stringValue(
        row,
        vehicleIndex,
      );

      if (vehicle.isEmpty) {
        continue;
      }

      result[vehicle] = {
        'بورسعيد': _numberValue(
              row,
              headers['محطة بورسعيد'],
            ) +
            _numberValue(
              row,
              headers['محطة بورفؤاد'],
            ),
        'إسماعيلية': _numberValue(
          row,
          headers['محطة الإسماعيلية'],
        ),
        'سويس': _numberValue(
          row,
          headers['محطة السويس'],
        ),
        'كارت ذكي': _numberValue(
          row,
          headers['تموين بالكارت الذكي'],
        ),
        'غاز': _numberValue(
          row,
          headers['غاز'],
        ),
      };
    }

    return result;
  }

  void _buildOutput(
    Sheet output,
    Sheet report,
    Map<String, Map<String, dynamic>> detailsMap,
  ) {
    const headers = [
      'م',
      'رقم السجل',
      'رقم السيارة',
      'الأحرف',
      'بورسعيد',
      'إسماعيلية',
      'سويس',
      'كارت ذكي',
      'غاز',
      'الكمية الإجمالية',
      'عداد البداية',
      'آخر عداد بالفترة',
      'المسافة',
      'النسبة',
      'النسبة القياسية',
      'نسبة التجاوز',
      'الحالة',
    ];

    output.merge(
      CellIndex.indexByString('A1'),
      CellIndex.indexByString('Q1'),
    );

    output.cell(
      CellIndex.indexByString('A1'),
    ).value = TextCellValue(
      'كشف نسبة السولار عن الشهر',
    );

    for (var i = 0; i < headers.length; i++) {
      output.cell(
        CellIndex.indexByColumnRow(
          columnIndex: i,
          rowIndex: 1,
        ),
      ).value = TextCellValue(headers[i]);
    }

    final rows = report.rows;

    if (rows.isEmpty) {
      return;
    }

    final reportHeaders = _findHeaders(rows);

    var outputRow = 2;
    var sequence = 1;

    for (var i = 1; i < rows.length; i++) {
      final row = rows[i];

      final vehicle = _stringValue(
        row,
        reportHeaders['رقم السيارة'],
      );

      if (vehicle.isEmpty) {
        continue;
      }

      final details = detailsMap[vehicle] ?? {};

      final excelRow = outputRow + 1;

      output
          .cell(
            CellIndex.indexByColumnRow(
              columnIndex: 0,
              rowIndex: outputRow,
            ),
          )
          .value = IntCellValue(sequence);

      _setValue(
        output,
        1,
        outputRow,
        _value(row, reportHeaders['رقم السجل']),
      );

      _setValue(
        output,
        2,
        outputRow,
        vehicle,
      );

      _setValue(
        output,
        3,
        outputRow,
        _value(row, reportHeaders['الأحرف']),
      );

      _setValue(
        output,
        4,
        outputRow,
        details['بورسعيد'] ?? 0,
      );

      _setValue(
        output,
        5,
        outputRow,
        details['إسماعيلية'] ?? 0,
      );

      _setValue(
        output,
        6,
        outputRow,
        details['سويس'] ?? 0,
      );

      _setValue(
        output,
        7,
        outputRow,
        details['كارت ذكي'] ?? 0,
      );

      _setValue(
        output,
        8,
        outputRow,
        details['غاز'] ?? 0,
      );

      output
          .cell(
            CellIndex.indexByColumnRow(
              columnIndex: 9,
              rowIndex: outputRow,
            ),
          )
          .value = FormulaCellValue(
        'SUM(E$excelRow:I$excelRow)',
      );

      _setValue(
        output,
        10,
        outputRow,
        _value(row, reportHeaders['عداد البداية']),
      );

      _setValue(
        output,
        11,
        outputRow,
        _value(row, reportHeaders['آخر عداد بالفترة']),
      );

      _setValue(
        output,
        12,
        outputRow,
        _value(row, reportHeaders['المسافة']),
      );

      _setValue(
        output,
        13,
        outputRow,
        _value(row, reportHeaders['النسبة']),
      );

      _setValue(
        output,
        14,
        outputRow,
        _value(row, reportHeaders['النسبة القياسية']),
      );

      _setValue(
        output,
        15,
        outputRow,
        _value(row, reportHeaders['نسبة التجاوز']),
      );

      _setValue(
        output,
        16,
        outputRow,
        _value(row, reportHeaders['الحالة']) ?? 'طبيعي',
      );

      sequence++;
      outputRow++;
    }
  }

  Map<String, int> _findHeaders(
    List<List<Data?>> rows,
  ) {
    if (rows.isEmpty) {
      return {};
    }

    final result = <String, int>{};

    for (var column = 0; column < rows.first.length; column++) {
      final value = rows.first[column]?.value?.toString().trim();

      if (value != null && value.isNotEmpty) {
        result[value] = column;
      }
    }

    return result;
  }

  dynamic _value(
    List<Data?> row,
    int? index,
  ) {
    if (index == null || index >= row.length) {
      return null;
    }

    return row[index]?.value;
  }

  String _stringValue(
    List<Data?> row,
    int? index,
  ) {
    final value = _value(row, index);

    if (value == null) {
      return '';
    }

    return value.toString().trim();
  }

  num _numberValue(
    List<Data?> row,
    int? index,
  ) {
    final value = _value(row, index);

    if (value == null) {
      return 0;
    }

    if (value is num) {
      return value;
    }

    return num.tryParse(
          value.toString().replaceAll(',', ''),
        ) ??
        0;
  }

  void _setValue(
    Sheet sheet,
    int column,
    int row,
    dynamic value,
  ) {
    final cell = sheet.cell(
      CellIndex.indexByColumnRow(
        columnIndex: column,
        rowIndex: row,
      ),
    );

    if (value == null) {
      return;
    }

    if (value is int) {
      cell.value = IntCellValue(value);
    } else if (value is double) {
      cell.value = DoubleCellValue(value);
    } else if (value is num) {
      cell.value = DoubleCellValue(value.toDouble());
    } else {
      cell.value = TextCellValue(value.toString());
    }
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

    html.AnchorElement(href: url)
      ..setAttribute('download', fileName)
      ..click();

    html.Url.revokeObjectUrl(url);
  }

  void _showError(String message) {
    if (!mounted) return;

    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(message),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('كشف النسبة'),
      ),
      body: Directionality(
        textDirection: TextDirection.rtl,
        child: Center(
          child: ConstrainedBox(
            constraints: const BoxConstraints(
              maxWidth: 600,
            ),
            child: Padding(
              padding: const EdgeInsets.all(24),
              child: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  const Icon(
                    Icons.table_chart_outlined,
                    size: 70,
                  ),
                  const SizedBox(height: 20),
                  const Text(
                    'كشف النسبة',
                    style: TextStyle(
                      fontSize: 24,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                  const SizedBox(height: 10),
                  Text(
                    _status,
                    textAlign: TextAlign.center,
                  ),
                  const SizedBox(height: 25),
                  if (_selectedFileName != null)
                    Text(
                      _selectedFileName!,
                      textAlign: TextAlign.center,
                    ),
                  const SizedBox(height: 20),
                  SizedBox(
                    width: double.infinity,
                    child: ElevatedButton.icon(
                      onPressed: _isProcessing ? null : _pickFile,
                      icon: const Icon(Icons.upload_file),
                      label: const Text('رفع ملف Excel'),
                    ),
                  ),
                  const SizedBox(height: 12),
                  SizedBox(
                    width: double.infinity,
                    child: ElevatedButton.icon(
                      onPressed: _isProcessing || _selectedBytes == null
                          ? null
                          : _processFile,
                      icon: _isProcessing
                          ? const SizedBox(
                              width: 18,
                              height: 18,
                              child: CircularProgressIndicator(
                                strokeWidth: 2,
                              ),
                            )
                          : const Icon(Icons.auto_fix_high),
                      label: Text(
                        _isProcessing
                            ? 'جاري المعالجة...'
                            : 'تنسيق وتصدير',
                      ),
                    ),
                  ),
                  if (_isReady) ...[
                    const SizedBox(height: 15),
                    const Text(
                      'تم إنشاء الملف وتصديره.',
                    ),
                  ],
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}
