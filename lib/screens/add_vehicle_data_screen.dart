import 'dart:typed_data';

import 'package:excel/excel.dart' as ex;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:universal_html/html.dart' as html;

/// حقل بيانات مركبة ممكن إضافته للتقرير كعمود جديد
class _VehicleField {
  final String label;
  bool selected = false;

  _VehicleField(this.label);
}

/// نتيجة تحديد صف العناوين داخل شيت (رقم الصف + خريطة اسم العمود -> رقمه)
class _ParsedHeader {
  final int headerRowIndex;
  final Map<String, int> headerMap;

  _ParsedHeader({
    required this.headerRowIndex,
    required this.headerMap,
  });
}

class AddVehicleDataScreen extends StatefulWidget {
  const AddVehicleDataScreen({super.key});

  @override
  State<AddVehicleDataScreen> createState() => _AddVehicleDataScreenState();
}

class _AddVehicleDataScreenState extends State<AddVehicleDataScreen> {
  Uint8List? _reportBytes;
  String? _reportFileName;

  Uint8List? _vehiclesBytes;
  String? _vehiclesFileName;

  bool _isProcessing = false;
  String _status = 'اختر ملف التقرير أولاً';

  Uint8List? _preparedFile;

  final TextEditingController _fileNameController =
      TextEditingController(text: 'التقرير المحدث');

  // نفس الترتيب دايمًا بغض النظر عن ترتيب اختيار المستخدم للـ checkboxes
  final List<_VehicleField> _fields = [
    _VehicleField('نوع السيارة'),
    _VehicleField('نوع الوقود'),
    _VehicleField('الماركة'),
    _VehicleField('سنة الصنع'),
  ];

  @override
  void dispose() {
    _fileNameController.dispose();
    super.dispose();
  }

  Future<void> _pickReportFile() async {
    final file = await FilePicker.pickFile(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (file == null) {
      return;
    }

    final bytes = await file.readAsBytes();

    setState(() {
      _reportBytes = bytes;
      _reportFileName = file.name;
      _preparedFile = null;
      _status = 'اختر ملف بيانات المركبات (لازم يحتوي على شيت "vehicles")';
    });
  }

  Future<void> _pickVehiclesFile() async {
    final file = await FilePicker.pickFile(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (file == null) {
      return;
    }

    final bytes = await file.readAsBytes();

    setState(() {
      _vehiclesBytes = bytes;
      _vehiclesFileName = file.name;
      _preparedFile = null;
      _status = 'اختر البيانات المطلوب إضافتها ثم اضغط "تجهيز التقرير"';
    });
  }

  bool get _canGenerate =>
      _reportBytes != null &&
      _vehiclesBytes != null &&
      _fields.any((f) => f.selected) &&
      !_isProcessing;

  Future<void> _generate() async {
    if (!_canGenerate) {
      return;
    }

    setState(() {
      _isProcessing = true;
      _preparedFile = null;
      _status = 'جاري تجهيز التقرير...';
    });

    try {
      final output = _buildUpdatedReport();

      if (!mounted) {
        return;
      }

      setState(() {
        _preparedFile = output;
        _isProcessing = false;
        _status = 'تم تجهيز التقرير بنجاح — اضغط تصدير الملف';
      });
    } catch (e) {
      if (!mounted) {
        return;
      }

      setState(() {
        _isProcessing = false;
        _preparedFile = null;
        _status = 'حدث خطأ: $e';
      });

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('خطأ: $e'),
          backgroundColor: Colors.red,
        ),
      );
    }
  }

  // =========================
  // المنطق الأساسي: قراءة التقرير + vehicles، المطابقة، إضافة الأعمدة
  // =========================
  Uint8List _buildUpdatedReport() {
    final reportExcel = ex.Excel.decodeBytes(_reportBytes!);
    final vehiclesExcel = ex.Excel.decodeBytes(_vehiclesBytes!);

    // ===== 1) إيجاد شيت "vehicles" بالاسم، بغض النظر عن ترتيبه =====
    ex.Sheet? vehiclesSheet;

    for (final entry in vehiclesExcel.tables.entries) {
      if (_normalize(entry.key).toLowerCase() == 'vehicles') {
        vehiclesSheet = entry.value;
        break;
      }
    }

    if (vehiclesSheet == null) {
      throw Exception(
        'لم يتم العثور على شيت باسم "vehicles" في ملف بيانات المركبات',
      );
    }

    final vehiclesHeaderInfo = _locateHeaderRow(
      vehiclesSheet,
      {'رقم السجل', 'رقم السيارة'},
    );

    if (vehiclesHeaderInfo == null) {
      throw Exception(
        'تعذر إيجاد صف عناوين في شيت "vehicles" (لازم يحتوي على '
        '"رقم السجل" أو "رقم السيارة")',
      );
    }

    final vehiclesHeaders = vehiclesHeaderInfo.headerMap;

    // ===== 2) إيجاد شيت التقرير وصف عناوينه (أي شيت، أي ترتيب أعمدة) =====
    ex.Sheet? reportSheet;
    int reportHeaderRow = -1;
    Map<String, int> reportHeaders = {};

    for (final table in reportExcel.tables.values) {
      final located = _locateHeaderRow(table, {'رقم السجل', 'رقم السيارة'});

      if (located != null) {
        reportSheet = table;
        reportHeaderRow = located.headerRowIndex;
        reportHeaders = located.headerMap;
        break;
      }
    }

    if (reportSheet == null) {
      throw Exception(
        'تعذر إيجاد صف عناوين في ملف التقرير (لازم يحتوي على '
        '"رقم السجل" أو "رقم السيارة")',
      );
    }

    // ===== 3) تحديد مفتاح المطابقة (رقم السجل أولًا، وإلا رقم السيارة) =====
    final reportHasRecord = reportHeaders.containsKey('رقم السجل');
    final vehiclesHasRecord = vehiclesHeaders.containsKey('رقم السجل');

    final useRecordKey = reportHasRecord && vehiclesHasRecord;
    final keyColumnName = useRecordKey ? 'رقم السجل' : 'رقم السيارة';

    final reportKeyIndex = reportHeaders[keyColumnName];
    final vehiclesKeyIndex = vehiclesHeaders[keyColumnName];

    if (reportKeyIndex == null || vehiclesKeyIndex == null) {
      throw Exception(
        'تعذرت المطابقة: عمود "$keyColumnName" لازم يكون موجود في '
        'كل من ملف التقرير وملف بيانات المركبات',
      );
    }

    // ===== 4) الحقول المختارة وفهارسها داخل vehicles =====
    final selectedFields = _fields.where((f) => f.selected).toList();

    final vehicleFieldIndex = <String, int?>{
      for (final field in selectedFields) field.label: vehiclesHeaders[field.label],
    };

    // ===== 5) بناء خريطة بيانات المركبات: المفتاح -> قيم الحقول =====
    final Map<String, Map<String, String>> vehiclesMap = {};
    final vRows = vehiclesSheet.rows;

    for (int r = vehiclesHeaderInfo.headerRowIndex + 1; r < vRows.length; r++) {
      final row = vRows[r];
      final key = _cellText(row, vehiclesKeyIndex).trim();

      if (key.isEmpty) {
        continue;
      }

      final values = <String, String>{};

      for (final field in selectedFields) {
        final idx = vehicleFieldIndex[field.label];
        values[field.label] = idx != null ? _cellText(row, idx).trim() : '';
      }

      vehiclesMap[key] = values;
    }

    // ===== 6) إضافة الأعمدة الجديدة في نهاية أعمدة التقرير الحالية =====
    final newStartColumn = reportSheet.maxColumns;

    // ننسخ تنسيق الهيدر من أول عمود موجود فعلًا في صف العناوين، عشان
    // عناوين الأعمدة الجديدة تطلع بنفس شكل هيدر التقرير بالظبط.
    final referenceHeaderStyle = reportSheet
        .cell(
          ex.CellIndex.indexByColumnRow(
            columnIndex: 0,
            rowIndex: reportHeaderRow,
          ),
        )
        .cellStyle;

    for (int i = 0; i < selectedFields.length; i++) {
      final cell = reportSheet.cell(
        ex.CellIndex.indexByColumnRow(
          columnIndex: newStartColumn + i,
          rowIndex: reportHeaderRow,
        ),
      );

      // مهم: القيمة الأول وبعدين التنسيق، وإلا الباكدج بيمسح الفورمات
      // وقت ما بيحدد صيغة الخلية تلقائيًا.
      cell.value = ex.TextCellValue(selectedFields[i].label);

      if (referenceHeaderStyle != null) {
        cell.cellStyle = referenceHeaderStyle;
      }

      reportSheet.setColumnWidth(newStartColumn + i, 16);
    }

    // ===== صفوف البيانات =====
    final rRows = reportSheet.rows;

    for (int r = reportHeaderRow + 1; r < rRows.length; r++) {
      final row = rRows[r];

      // ننسخ تنسيق نفس الصف (بما فيه لون الصف المتبادل لو موجود) من
      // أول عمود موجود فعلًا في الصف ده، عشان الأعمدة الجديدة تندمج
      // بصريًا بنفس شكل باقي التقرير.
      final referenceRowStyle = reportSheet
          .cell(
            ex.CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: r),
          )
          .cellStyle;

      final key = _cellText(row, reportKeyIndex).trim();
      final matchedValues = key.isEmpty ? null : vehiclesMap[key];

      for (int i = 0; i < selectedFields.length; i++) {
        final field = selectedFields[i];
        final value = matchedValues?[field.label] ?? '';

        final cell = reportSheet.cell(
          ex.CellIndex.indexByColumnRow(
            columnIndex: newStartColumn + i,
            rowIndex: r,
          ),
        );

        if (value.isNotEmpty) {
          cell.value = ex.TextCellValue(value);
        }

        if (referenceRowStyle != null) {
          cell.cellStyle = referenceRowStyle;
        }
      }
    }

    // نتأكد إن الشيت لسه من اليمين لليسار زي الأصل
    reportSheet.isRTL = true;

    final encoded = reportExcel.encode();

    if (encoded == null) {
      throw Exception('تعذر إنشاء ملف Excel النهائي');
    }

    return Uint8List.fromList(encoded);
  }

  /// بيدوّر عن أول صف (في أول 10 صفوف) فيه أي عمود من أسماء المفاتيح
  /// المطلوبة، ويرجّع خريطة "اسم العمود -> رقمه" لكل عناوين الصف ده،
  /// بغض النظر عن ترتيب الأعمدة أو شكل الشيت.
  _ParsedHeader? _locateHeaderRow(ex.Sheet sheet, Set<String> anyOfKeys) {
    final rows = sheet.rows;
    final maxScan = rows.length < 10 ? rows.length : 10;

    for (int r = 0; r < maxScan; r++) {
      final headerMap = _findHeaders(rows[r]);

      for (final key in anyOfKeys) {
        if (headerMap.containsKey(key)) {
          return _ParsedHeader(headerRowIndex: r, headerMap: headerMap);
        }
      }
    }

    return null;
  }

  Map<String, int> _findHeaders(List<ex.Data?> row) {
    final result = <String, int>{};

    for (int column = 0; column < row.length; column++) {
      final text = _normalize(_cellText(row, column));

      if (text.isNotEmpty) {
        result[text] = column;
      }
    }

    return result;
  }

  /// إزالة المسافات الزايدة (بما فيها non-breaking space) من اسم
  /// العمود بس، من غير ما نلمس محتوى البيانات نفسها.
  String _normalize(String value) {
    return value
        .replaceAll('\u00A0', ' ')
        .replaceAll(RegExp(r'\s+'), ' ')
        .trim();
  }

  String _cellText(List<ex.Data?> row, int? index) {
    if (index == null || index < 0 || index >= row.length) {
      return '';
    }

    final data = row[index];

    if (data == null || data.value == null) {
      return '';
    }

    final value = data.value!;

    if (value is ex.TextCellValue) {
      return value.value.text ?? '';
    }

    if (value is ex.IntCellValue) {
      return value.value.toString();
    }

    if (value is ex.DoubleCellValue) {
      return value.value.toString();
    }

    if (value is ex.FormulaCellValue) {
      return '';
    }

    return value.toString();
  }

  void _exportFile() {
    if (_preparedFile == null) {
      return;
    }

    String fileName = _fileNameController.text.trim();

    if (fileName.isEmpty) {
      fileName = 'التقرير المحدث';
    }

    if (!fileName.toLowerCase().endsWith('.xlsx')) {
      fileName = '$fileName.xlsx';
    }

    final blob = html.Blob(
      [_preparedFile!],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    final url = html.Url.createObjectUrlFromBlob(blob);

    html.AnchorElement(href: url)
      ..setAttribute('download', fileName)
      ..click();

    html.Url.revokeObjectUrl(url);

    setState(() {
      _status = 'تم تصدير الملف بنجاح';
    });
  }

  // =========================
  // الواجهة
  // =========================
  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: TextDirection.rtl,
      child: Scaffold(
        appBar: AppBar(
          title: const Text('إضافة بيانات للتقارير'),
          centerTitle: true,
        ),
        body: Center(
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 650),
            child: SingleChildScrollView(
              padding: const EdgeInsets.all(24),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  const Icon(
                    Icons.playlist_add_rounded,
                    size: 64,
                  ),
                  const SizedBox(height: 16),
                  const Text(
                    'إضافة بيانات للتقارير',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 22,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                  const SizedBox(height: 24),

                  // ===== 1) اختيار ملف التقرير =====
                  Card(
                    child: ListTile(
                      leading: const Icon(Icons.description_outlined),
                      title: const Text('١) ملف التقرير'),
                      subtitle: Text(
                        _reportFileName ?? 'لم يتم اختيار ملف بعد',
                        style: TextStyle(
                          color: _reportFileName != null
                              ? Colors.green.shade700
                              : Colors.grey.shade600,
                        ),
                      ),
                      trailing: IconButton(
                        icon: Icon(
                          _reportFileName != null
                              ? Icons.check_circle
                              : Icons.upload_file,
                          color: _reportFileName != null ? Colors.green : null,
                        ),
                        onPressed: _pickReportFile,
                      ),
                      onTap: _pickReportFile,
                    ),
                  ),

                  const SizedBox(height: 12),

                  // ===== 2) اختيار ملف بيانات المركبات =====
                  Card(
                    child: ListTile(
                      leading: const Icon(Icons.directions_car_filled_outlined),
                      title: const Text('٢) ملف بيانات المركبات (vehicles)'),
                      subtitle: Text(
                        _vehiclesFileName ?? 'لم يتم اختيار ملف بعد',
                        style: TextStyle(
                          color: _vehiclesFileName != null
                              ? Colors.green.shade700
                              : Colors.grey.shade600,
                        ),
                      ),
                      trailing: IconButton(
                        icon: Icon(
                          _vehiclesFileName != null
                              ? Icons.check_circle
                              : Icons.upload_file,
                          color:
                              _vehiclesFileName != null ? Colors.green : null,
                        ),
                        onPressed:
                            _reportBytes == null ? null : _pickVehiclesFile,
                      ),
                      onTap: _reportBytes == null ? null : _pickVehiclesFile,
                    ),
                  ),

                  const SizedBox(height: 16),

                  // ===== 3) اختيار البيانات المراد إضافتها =====
                  if (_vehiclesBytes != null) ...[
                    const Align(
                      alignment: Alignment.centerRight,
                      child: Text(
                        '٣) البيانات المراد إضافتها:',
                        style: TextStyle(fontWeight: FontWeight.bold),
                      ),
                    ),
                    const SizedBox(height: 4),
                    Card(
                      child: Column(
                        children: [
                          for (final field in _fields)
                            CheckboxListTile(
                              value: field.selected,
                              title: Text(field.label),
                              controlAffinity:
                                  ListTileControlAffinity.leading,
                              onChanged: (checked) {
                                setState(() {
                                  field.selected = checked ?? false;
                                  _preparedFile = null;
                                });
                              },
                            ),
                        ],
                      ),
                    ),
                    const SizedBox(height: 16),
                  ],

                  // اسم الملف
                  TextField(
                    controller: _fileNameController,
                    textDirection: TextDirection.rtl,
                    decoration: const InputDecoration(
                      labelText: 'اسم الملف',
                      border: OutlineInputBorder(),
                    ),
                  ),

                  const SizedBox(height: 16),

                  Text(
                    _status,
                    textAlign: TextAlign.center,
                  ),

                  const SizedBox(height: 16),

                  ElevatedButton.icon(
                    onPressed: _canGenerate ? _generate : null,
                    icon: _isProcessing
                        ? const SizedBox(
                            width: 20,
                            height: 20,
                            child: CircularProgressIndicator(strokeWidth: 2),
                          )
                        : const Icon(Icons.build_outlined),
                    label: Text(
                      _isProcessing ? 'جاري التجهيز...' : 'تجهيز التقرير',
                    ),
                  ),

                  const SizedBox(height: 12),

                  FilledButton.icon(
                    onPressed: _preparedFile == null ? null : _exportFile,
                    icon: const Icon(Icons.download),
                    label: const Text('تصدير الملف'),
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
