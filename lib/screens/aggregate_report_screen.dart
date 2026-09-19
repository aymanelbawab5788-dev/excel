import 'dart:typed_data';

import 'package:excel/excel.dart' as ex;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:universal_html/html.dart' as html;

/// أسماء الشهور بالعربي، الفهرس 0 = يناير ... 11 = ديسمبر
const List<String> _kArabicMonthNames = [
  'يناير',
  'فبراير',
  'مارس',
  'أبريل',
  'مايو',
  'يونيو',
  'يوليو',
  'أغسطس',
  'سبتمبر',
  'أكتوبر',
  'نوفمبر',
  'ديسمبر',
];

/// خانة شهر واحد داخل الفترة المختارة
class _MonthSlot {
  final int monthNumber;
  Uint8List? fileBytes;
  String? fileName;

  _MonthSlot(this.monthNumber);

  String get monthName => _kArabicMonthNames[monthNumber - 1];
}

/// بيانات سيارة واحدة متجمعة عبر كل شهور الفترة
class _CarAggregate {
  final String recordNumber;
  String carNumber;
  String letters;
  num? startOdometer;
  num? endOdometer;

  /// المفتاح = فهرس الشهر داخل الفترة
  final Map<int, num> monthlyQuantities = {};

  _CarAggregate({
    required this.recordNumber,
    required this.carNumber,
    required this.letters,
  });
}

class AggregateReportScreen extends StatefulWidget {
  const AggregateReportScreen({super.key});

  @override
  State<AggregateReportScreen> createState() =>
      _AggregateReportScreenState();
}

class _AggregateReportScreenState extends State<AggregateReportScreen> {
  int? _fromMonth;
  int? _toMonth;

  List<_MonthSlot> _monthSlots = [];

  bool _isProcessing = false;
  String _status = 'اختر الفترة (من شهر - إلى شهر) أولاً';

  Uint8List? _preparedFile;

  final TextEditingController _fileNameController =
      TextEditingController(text: 'التقرير التجميعي');

  // ===== ألوان التنسيق =====
  static const String _kTitleColor = 'FF000000';
  static const String _kHeaderColor = 'FF2E5395';
  static const String _kHeaderFontColor = 'FFFFFFFF';
  static const String _kAltRowColor = 'FFF2F2F2';
  static const String _kBorderColor = 'FFBFBFBF';
  static const String _kOdometerColor = 'FFE2EFDA';
  static const String _kTotalColor = 'FFD9E2F3';

  @override
  void dispose() {
    _fileNameController.dispose();
    super.dispose();
  }

  // =========================
  // بناء تسلسل الشهور
  // =========================

  List<int> _buildMonthSequence(int from, int to) {
    final sequence = <int>[];
    int current = from;

    while (true) {
      sequence.add(current);

      if (current == to) {
        break;
      }

      current = current == 12 ? 1 : current + 1;

      if (sequence.length > 12) {
        break;
      }
    }

    return sequence;
  }

  void _applyPeriod() {
    if (_fromMonth == null || _toMonth == null) {
      return;
    }

    final sequence = _buildMonthSequence(
      _fromMonth!,
      _toMonth!,
    );

    setState(() {
      _monthSlots = sequence.map((m) => _MonthSlot(m)).toList();
      _preparedFile = null;
      _status = 'اختر ملف كل شهر من الشهور اللي ظهرت تحت';
    });
  }

  Future<void> _pickFileForSlot(_MonthSlot slot) async {
    final file = await FilePicker.pickFile(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (file == null) {
      return;
    }

    final bytes = await file.readAsBytes();

    setState(() {
      slot.fileBytes = bytes;
      slot.fileName = file.name;
      _preparedFile = null;
    });
  }

  bool get _allSlotsFilled =>
      _monthSlots.isNotEmpty &&
      _monthSlots.every((slot) => slot.fileBytes != null);

  Future<void> _generateReport() async {
    if (!_allSlotsFilled) {
      return;
    }

    setState(() {
      _isProcessing = true;
      _preparedFile = null;
      _status = 'جاري تجميع التقرير...';
    });

    try {
      final output = _buildAggregateReport(_monthSlots);

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
    }
  }

  // =========================
  // قراءة وتجميع الملفات
  // =========================

  Uint8List _buildAggregateReport(List<_MonthSlot> slots) {
    final Map<String, _CarAggregate> carsByRecord = {};
    final List<String> recordOrder = [];

    for (int slotIndex = 0; slotIndex < slots.length; slotIndex++) {
      final slot = slots[slotIndex];
      final bytes = slot.fileBytes!;

      final monthlyExcel = ex.Excel.decodeBytes(bytes);

      final parsed = _locateHeaderRow(monthlyExcel);

      if (parsed == null) {
        throw Exception(
          'تعذر إيجاد صف عناوين فيه "رقم السجل" في ملف شهر ${slot.monthName}',
        );
      }

      final dataSheet = parsed.sheet;
      final headerRowIndex = parsed.headerRowIndex;
      final headerMap = parsed.headerMap;

      final recordIndex = headerMap['رقم السجل'];

      if (recordIndex == null) {
        throw Exception(
          'عمود "رقم السجل" غير موجود في ملف شهر ${slot.monthName}',
        );
      }

      final carNumIndex = headerMap['رقم السيارة'];
      final lettersIndex = headerMap['الأحرف'];

      final portSaidIndex = headerMap['بورسعيد'];
      final ismailiaIndex = headerMap['إسماعيلية'];
      final suezIndex = headerMap['سويس'];
      final smartCardIndex = headerMap['كارت ذكي'];
      final gasIndex = headerMap['غاز'];

      final startOdoIndex = headerMap['عداد البداية'];
      final endOdoIndex = headerMap['آخر عداد بالفترة'];

      final rows = dataSheet.rows;

      for (int r = headerRowIndex + 1; r < rows.length; r++) {
        final row = rows[r];

        final recordNumber = _cellText(
          row,
          recordIndex,
        ).trim();

        if (recordNumber.isEmpty || recordNumber == 'الإجمالي') {
          continue;
        }

        final carNumber = carNumIndex != null
            ? _cellText(row, carNumIndex).trim()
            : '';

        final letters = lettersIndex != null
            ? _cellText(row, lettersIndex).trim()
            : '';

        // إعادة حساب كمية الشهر من الأعمدة الأصلية
        // بدل الاعتماد على قيمة FormulaCellValue.
        final quantity = _numberValue(row, portSaidIndex) +
            _numberValue(row, ismailiaIndex) +
            _numberValue(row, suezIndex) +
            _numberValue(row, smartCardIndex) +
            _numberValue(row, gasIndex);

        final startOdo = _numberValue(
          row,
          startOdoIndex,
        );

        final endOdo = _numberValue(
          row,
          endOdoIndex,
        );

        final isNewCar =
            !carsByRecord.containsKey(recordNumber);

        final car = carsByRecord.putIfAbsent(
          recordNumber,
          () => _CarAggregate(
            recordNumber: recordNumber,
            carNumber: carNumber,
            letters: letters,
          ),
        );

        if (isNewCar) {
          recordOrder.add(recordNumber);
        }

        // عداد بداية الفترة:
        // أول شهر تظهر فيه السيارة فعليًا.
        car.startOdometer ??= startOdo;

        // عداد نهاية الفترة:
        // آخر شهر تظهر فيه السيارة.
        car.endOdometer = endOdo;

        // كمية هذا الشهر.
        car.monthlyQuantities[slotIndex] = quantity;

        if (car.carNumber.isEmpty && carNumber.isNotEmpty) {
          car.carNumber = carNumber;
        }

        if (car.letters.isEmpty && letters.isNotEmpty) {
          car.letters = letters;
        }
      }
    }

    return _writeOutput(
      slots,
      carsByRecord,
      recordOrder,
    );
  }

  // =========================
  // البحث عن صف العناوين
  // =========================

  _ParsedHeader? _locateHeaderRow(ex.Excel excel) {
    for (final table in excel.tables.values) {
      final rows = table.rows;

      final maxRowsToScan =
          rows.length < 5 ? rows.length : 5;

      for (int r = 0; r < maxRowsToScan; r++) {
        final headerMap = _findHeaders(rows[r]);

        if (headerMap.containsKey('رقم السجل')) {
          return _ParsedHeader(
            sheet: table,
            headerRowIndex: r,
            headerMap: headerMap,
          );
        }
      }
    }

    return null;
  }

  Map<String, int> _findHeaders(List<ex.Data?> row) {
    final result = <String, int>{};

    for (int column = 0; column < row.length; column++) {
      final text = _cellText(
        row,
        column,
      ).trim();

      if (text.isNotEmpty) {
        result[text] = column;
      }
    }

    return result;
  }

  // =========================
  // قراءة الخلية كنص
  // =========================

  String _cellText(
    List<ex.Data?> row,
    int? index,
  ) {
    if (index == null ||
        index < 0 ||
        index >= row.length) {
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

    return value.toString();
  }

  // =========================
  // قراءة الخلية كرقم
  // =========================

  num _numberValue(
    List<ex.Data?> row,
    int? index,
  ) {
    if (index == null ||
        index < 0 ||
        index >= row.length) {
      return 0;
    }

    final data = row[index];

    if (data == null || data.value == null) {
      return 0;
    }

    final value = data.value!;

    if (value is ex.IntCellValue) {
      return value.value;
    }

    if (value is ex.DoubleCellValue) {
      return value.value;
    }

    if (value is ex.FormulaCellValue) {
      return 0;
    }

    if (value is ex.TextCellValue) {
      final raw = value.value.text ?? '';
      final cleaned = raw.replaceAll(',', '').trim();

      return num.tryParse(cleaned) ?? 0;
    }

    return num.tryParse(value.toString()) ?? 0;
  }

  // =========================
  // بناء ملف التقرير النهائي
  // =========================

  Uint8List _writeOutput(
    List<_MonthSlot> slots,
    Map<String, _CarAggregate> carsByRecord,
    List<String> recordOrder,
  ) {
    final outputExcel = ex.Excel.createExcel();

    final sheet = outputExcel['التقرير التجميعي'];

    final defaultSheet = outputExcel.getDefaultSheet();

    if (defaultSheet != null &&
        defaultSheet != 'التقرير التجميعي') {
      outputExcel.delete(defaultSheet);
    }

    sheet.isRTL = true;

    // ==================================================
    // ترتيب الأعمدة النهائي
    //
    // م
    // رقم السجل
    // رقم السيارة
    // الأحرف
    // يناير
    // فبراير
    // ...
    // إجمالي الكمية
    // عداد بداية الفترة
    // عداد نهاية الفترة
    // إجمالي المسافة
    // ==================================================

    final monthCount = slots.length;

    final firstMonthColumn = 4;

    final quantityTotalColumn =
        firstMonthColumn + monthCount;

    final startOdoColumn =
        quantityTotalColumn + 1;

    final endOdoColumn =
        startOdoColumn + 1;

    final periodTotalColumn =
        endOdoColumn + 1;

    final lastColumn = periodTotalColumn;

    // =========================
    // رؤوس الأعمدة
    // =========================

    final headers = <String>[
      'م',
      'رقم السجل',
      'رقم السيارة',
      'الأحرف',
      for (final slot in slots) slot.monthName,
      'إجمالي الكمية',
      'عداد بداية الفترة',
      'عداد نهاية الفترة',
      'إجمالي المسافة',
    ];

    // =========================
    // عنوان التقرير
    // =========================

    final periodLabel = slots.length == 1
        ? slots.first.monthName
        : '${slots.first.monthName} إلى ${slots.last.monthName}';

    sheet.merge(
      ex.CellIndex.indexByColumnRow(
        columnIndex: 0,
        rowIndex: 0,
      ),
      ex.CellIndex.indexByColumnRow(
        columnIndex: lastColumn,
        rowIndex: 0,
      ),
    );

    _writeCell(
      sheet,
      0,
      0,
      'التقرير التجميعي - $periodLabel',
      backgroundHex: _kTitleColor,
      fontColorHex: _kHeaderFontColor,
      bold: true,
      fontSize: 16,
    );

    sheet.setRowHeight(0, 28);

    // =========================
    // صف رؤوس الأعمدة
    // =========================

    for (int c = 0; c < headers.length; c++) {
      _writeCell(
        sheet,
        c,
        1,
        headers[c],
        backgroundHex: _kHeaderColor,
        fontColorHex: _kHeaderFontColor,
        bold: true,
      );
    }

    sheet.setRowHeight(1, 22);

    // =========================
    // صفوف السيارات
    // =========================

    var outputRow = 2;
    var serial = 1;

    for (final recordNumber in recordOrder) {
      final car = carsByRecord[recordNumber]!;

      final isAltRow = serial.isOdd;

      // م
      _writeCell(
        sheet,
        0,
        outputRow,
        serial,
        isAltRow: isAltRow,
      );

      // رقم السجل
      _writeCell(
        sheet,
        1,
        outputRow,
        car.recordNumber,
        isAltRow: isAltRow,
      );

      // رقم السيارة
      _writeCell(
        sheet,
        2,
        outputRow,
        car.carNumber,
        isAltRow: isAltRow,
      );

      // الأحرف
      _writeCell(
        sheet,
        3,
        outputRow,
        car.letters,
        isAltRow: isAltRow,
      );

      // =========================
      // الشهور
      // =========================

      for (int m = 0; m < monthCount; m++) {
        final quantity =
            car.monthlyQuantities[m] ?? 0;

        _writeCell(
          sheet,
          firstMonthColumn + m,
          outputRow,
          quantity,
          isAltRow: isAltRow,
        );
      }

      final excelRow = outputRow + 1;

      // =========================
      // إجمالي الكمية
      // =========================

      final fromMonthCol =
          _columnLetter(firstMonthColumn);

      final toMonthCol =
          _columnLetter(
        firstMonthColumn + monthCount - 1,
      );

      sheet.updateCell(
        ex.CellIndex.indexByColumnRow(
          columnIndex: quantityTotalColumn,
          rowIndex: outputRow,
        ),
        ex.FormulaCellValue(
          'SUM($fromMonthCol$excelRow:$toMonthCol$excelRow)',
        ),
        cellStyle: _cellStyle(
          backgroundHex: _kTotalColor,
          bold: true,
        ),
      );

      // =========================
      // عداد بداية الفترة
      // =========================

      _writeCell(
        sheet,
        startOdoColumn,
        outputRow,
        car.startOdometer ?? 0,
        backgroundHex: _kOdometerColor,
      );

      // =========================
      // عداد نهاية الفترة
      // =========================

      _writeCell(
        sheet,
        endOdoColumn,
        outputRow,
        car.endOdometer ?? 0,
        backgroundHex: _kOdometerColor,
      );

      // =========================
      // إجمالي المسافة
      // = عداد النهاية - عداد البداية
      // =========================

      final startOdoColLetter =
          _columnLetter(startOdoColumn);

      final endOdoColLetter =
          _columnLetter(endOdoColumn);

      sheet.updateCell(
        ex.CellIndex.indexByColumnRow(
          columnIndex: periodTotalColumn,
          rowIndex: outputRow,
        ),
        ex.FormulaCellValue(
          '$endOdoColLetter$excelRow-$startOdoColLetter$excelRow',
        ),
        cellStyle: _cellStyle(
          backgroundHex: _kTotalColor,
          bold: true,
        ),
      );

      outputRow++;
      serial++;
    }

    // =========================
    // عرض الأعمدة
    // =========================

    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 12);
    sheet.setColumnWidth(2, 14);
    sheet.setColumnWidth(3, 10);

    for (int m = 0; m < monthCount; m++) {
      sheet.setColumnWidth(
        firstMonthColumn + m,
        12,
      );
    }

    sheet.setColumnWidth(
      quantityTotalColumn,
      14,
    );

    sheet.setColumnWidth(
      startOdoColumn,
      14,
    );

    sheet.setColumnWidth(
      endOdoColumn,
      14,
    );

    sheet.setColumnWidth(
      periodTotalColumn,
      14,
    );

    // =========================
    // إنشاء الملف
    // =========================

    final encoded = outputExcel.encode();

    if (encoded == null) {
      throw Exception(
        'فشل إنشاء ملف التقرير التجميعي',
      );
    }

    return Uint8List.fromList(encoded);
  }

  // =========================
  // تحويل رقم العمود إلى حرف Excel
  // =========================

  String _columnLetter(int index) {
    int number = index + 1;

    String result = '';

    while (number > 0) {
      final remainder = (number - 1) % 26;

      result =
          String.fromCharCode(65 + remainder) + result;

      number = (number - 1) ~/ 26;
    }

    return result;
  }

  // =========================
  // تنسيق الخلايا
  // =========================

  ex.CellStyle _cellStyle({
    String backgroundHex = 'FFFFFFFF',
    String fontColorHex = 'FF000000',
    bool bold = false,
    int fontSize = 11,
  }) {
    final border = ex.Border(
      borderStyle: ex.BorderStyle.Thin,
      borderColorHex:
          ex.ExcelColor.fromHexString(
        _kBorderColor,
      ),
    );

    return ex.CellStyle(
      backgroundColorHex:
          ex.ExcelColor.fromHexString(
        backgroundHex,
      ),
      fontColorHex:
          ex.ExcelColor.fromHexString(
        fontColorHex,
      ),
      bold: bold,
      fontSize: fontSize,
      horizontalAlign:
          ex.HorizontalAlign.Center,
      verticalAlign:
          ex.VerticalAlign.Center,
      leftBorder: border,
      rightBorder: border,
      topBorder: border,
      bottomBorder: border,
    );
  }

  // =========================
  // كتابة الخلية
  // =========================

  void _writeCell(
    ex.Sheet sheet,
    int column,
    int row,
    dynamic value, {
    bool isAltRow = false,
    String? backgroundHex,
    String fontColorHex = 'FF000000',
    bool bold = false,
    int fontSize = 11,
  }) {
    final cell = sheet.cell(
      ex.CellIndex.indexByColumnRow(
        columnIndex: column,
        rowIndex: row,
      ),
    );

    if (value is int) {
      cell.value = ex.IntCellValue(value);
    } else if (value is double) {
      cell.value = ex.DoubleCellValue(value);
    } else if (value is num) {
      cell.value =
          ex.DoubleCellValue(value.toDouble());
    } else {
      cell.value =
          ex.TextCellValue(value.toString());
    }

    final resolvedBackground =
        backgroundHex ??
            (isAltRow
                ? _kAltRowColor
                : 'FFFFFFFF');

    cell.cellStyle = _cellStyle(
      backgroundHex: resolvedBackground,
      fontColorHex: fontColorHex,
      bold: bold,
      fontSize: fontSize,
    );
  }

  // =========================
  // تصدير الملف
  // =========================

  void _exportFile() {
    if (_preparedFile == null) {
      return;
    }

    String fileName =
        _fileNameController.text.trim();

    if (fileName.isEmpty) {
      fileName = 'التقرير التجميعي';
    }

    if (!fileName
        .toLowerCase()
        .endsWith('.xlsx')) {
      fileName = '$fileName.xlsx';
    }

    final blob = html.Blob(
      [_preparedFile!],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    final url =
        html.Url.createObjectUrlFromBlob(blob);

    html.AnchorElement(href: url)
      ..setAttribute(
        'download',
        fileName,
      )
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
          title:
              const Text('التقرير التجميعي'),
          centerTitle: true,
        ),
        body: Center(
          child: ConstrainedBox(
            constraints:
                const BoxConstraints(
              maxWidth: 650,
            ),
            child:
                SingleChildScrollView(
              padding:
                  const EdgeInsets.all(24),
              child: Column(
                crossAxisAlignment:
                    CrossAxisAlignment.stretch,
                children: [
                  const Icon(
                    Icons.summarize_outlined,
                    size: 64,
                  ),
                  const SizedBox(
                    height: 16,
                  ),
                  const Text(
                    'التقرير التجميعي',
                    textAlign:
                        TextAlign.center,
                    style: TextStyle(
                      fontSize: 22,
                      fontWeight:
                          FontWeight.bold,
                    ),
                  ),
                  const SizedBox(
                    height: 24,
                  ),

                  // اختيار الفترة
                  Row(
                    children: [
                      Expanded(
                        child:
                            DropdownButtonFormField<
                                int>(
                          initialValue:
                              _fromMonth,
                          decoration:
                              const InputDecoration(
                            labelText:
                                'من شهر',
                            border:
                                OutlineInputBorder(),
                          ),
                          items: [
                            for (int m = 1;
                                m <= 12;
                                m++)
                              DropdownMenuItem(
                                value: m,
                                child: Text(
                                  _kArabicMonthNames[
                                      m - 1],
                                ),
                              ),
                          ],
                          onChanged: (value) {
                            setState(() {
                              _fromMonth =
                                  value;
                            });
                          },
                        ),
                      ),
                      const SizedBox(
                        width: 12,
                      ),
                      Expanded(
                        child:
                            DropdownButtonFormField<
                                int>(
                          initialValue:
                              _toMonth,
                          decoration:
                              const InputDecoration(
                            labelText:
                                'إلى شهر',
                            border:
                                OutlineInputBorder(),
                          ),
                          items: [
                            for (int m = 1;
                                m <= 12;
                                m++)
                              DropdownMenuItem(
                                value: m,
                                child: Text(
                                  _kArabicMonthNames[
                                      m - 1],
                                ),
                              ),
                          ],
                          onChanged: (value) {
                            setState(() {
                              _toMonth =
                                  value;
                            });
                          },
                        ),
                      ),
                    ],
                  ),

                  const SizedBox(
                    height: 12,
                  ),

                  ElevatedButton(
                    onPressed:
                        (_fromMonth == null ||
                                _toMonth == null)
                            ? null
                            : _applyPeriod,
                    child: const Text(
                      'تحديد الفترة',
                    ),
                  ),

                  const SizedBox(
                    height: 20,
                  ),

                  // ملفات الشهور
                  if (_monthSlots.isNotEmpty) ...[
                    const Align(
                      alignment:
                          Alignment.centerRight,
                      child: Text(
                        'ارفع ملف كل شهر:',
                        style: TextStyle(
                          fontWeight:
                              FontWeight.bold,
                        ),
                      ),
                    ),
                    const SizedBox(
                      height: 8,
                    ),
                    for (final slot
                        in _monthSlots)
                      Card(
                        child: ListTile(
                          title: Text(
                            slot.monthName,
                          ),
                          subtitle:
                              Text(
                            slot.fileName ??
                                'لم يتم اختيار ملف بعد',
                            style:
                                TextStyle(
                              color: slot
                                      .fileBytes !=
                                  null
                                  ? Colors
                                      .green
                                      .shade700
                                  : Colors
                                      .grey
                                      .shade600,
                            ),
                          ),
                          trailing:
                              IconButton(
                            icon: Icon(
                              slot.fileBytes !=
                                      null
                                  ? Icons
                                      .check_circle
                                  : Icons
                                      .upload_file,
                              color: slot
                                      .fileBytes !=
                                  null
                                  ? Colors
                                      .green
                                  : null,
                            ),
                            onPressed: () =>
                                _pickFileForSlot(
                              slot,
                            ),
                          ),
                          onTap: () =>
                              _pickFileForSlot(
                            slot,
                          ),
                        ),
                      ),
                    const SizedBox(
                      height: 16,
                    ),
                  ],

                  // اسم الملف
                  TextField(
                    controller:
                        _fileNameController,
                    textDirection:
                        TextDirection.rtl,
                    decoration:
                        const InputDecoration(
                      labelText:
                          'اسم الملف',
                      border:
                          OutlineInputBorder(),
                    ),
                  ),

                  const SizedBox(
                    height: 16,
                  ),

                  Text(
                    _status,
                    textAlign:
                        TextAlign.center,
                  ),

                  const SizedBox(
                    height: 16,
                  ),

                  ElevatedButton.icon(
                    onPressed:
                        (!_allSlotsFilled ||
                                _isProcessing)
                            ? null
                            : _generateReport,
                    icon: _isProcessing
                        ? const SizedBox(
                            width: 20,
                            height: 20,
                            child:
                                CircularProgressIndicator(
                              strokeWidth: 2,
                            ),
                          )
                        : const Icon(
                            Icons.summarize,
                          ),
                    label: Text(
                      _isProcessing
                          ? 'جاري التجميع...'
                          : 'توليد التقرير',
                    ),
                  ),

                  const SizedBox(
                    height: 12,
                  ),

                  FilledButton.icon(
                    onPressed:
                        _preparedFile == null
                            ? null
                            : _exportFile,
                    icon: const Icon(
                      Icons.download,
                    ),
                    label: const Text(
                      'تصدير الملف',
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

class _ParsedHeader {
  final ex.Sheet sheet;
  final int headerRowIndex;
  final Map<String, int> headerMap;

  _ParsedHeader({
    required this.sheet,
    required this.headerRowIndex,
    required this.headerMap,
  });
}
