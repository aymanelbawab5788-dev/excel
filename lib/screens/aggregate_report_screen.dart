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

/// خانة شهر واحد داخل الفترة المختارة: رقم الشهر + الملف اللي هيتربط بيه
class _MonthSlot {
  final int monthNumber; // 1..12
  Uint8List? fileBytes;
  String? fileName;

  _MonthSlot(this.monthNumber);

  String get monthName => _kArabicMonthNames[monthNumber - 1];
}

/// بيانات سيارة واحدة متجمّعة عبر كل شهور الفترة
class _CarAggregate {
  final String recordNumber; // المفتاح الأساسي
  String carNumber;
  String letters;
  num? startOdometer;
  num? endOdometer;

  /// key = فهرس الشهر داخل قائمة الفترة (0 = أول شهر بالفترة)
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
  // بناء تسلسل الشهور الزمني (بيتعامل مع عبور نهاية السنة)
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

      // حماية من حلقة لا نهائية لو حصل خطأ غير متوقع
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

    final sequence = _buildMonthSequence(_fromMonth!, _toMonth!);

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
  // المنطق الأساسي: قراءة كل شهر وتجميع بيانات السيارات
  // =========================
  Uint8List _buildAggregateReport(List<_MonthSlot> slots) {
    // key = رقم السجل
    final Map<String, _CarAggregate> carsByRecord = {};

    // بنحافظ على ترتيب أول ظهور لكل سيارة عشان نرتب صفوف التقرير بنفس
    // منطق "أول ظهور بالفترة"، مش بترتيب عشوائي
    final List<String> recordOrder = [];

    // مهم جدًا: بنمرّ على الشهور بترتيبها الزمني في "slots" (اللي هي
    // أصلاً مبنية من _buildMonthSequence)، مش بأي ترتيب تاني - وده
    // اللي بيضمن إن عداد البداية والنهاية يتحسبوا صح حتى لو المستخدم
    // اختار/رفع الملفات مش بترتيبها.
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

        final recordNumber = _cellText(row, recordIndex).trim();

        // تجاهل الصفوف الفاضية وصف "الإجمالي" في آخر الجدول
        if (recordNumber.isEmpty || recordNumber == 'الإجمالي') {
          continue;
        }

        final carNumber =
            carNumIndex != null ? _cellText(row, carNumIndex).trim() : '';
        final letters =
            lettersIndex != null ? _cellText(row, lettersIndex).trim() : '';

        // ===== الكمية الإجمالية =====
        // ملحوظة مهمة: عمود "الكمية الإجمالية" في ملف كشف النسبة
        // مكتوب كمعادلة إكسل (=SUM(...))، ومعادلة من غير قيمة محفوظة
        // (لو الملف اتصدّر من البرنامج ومتفتحش في إكسل حقيقي قبل كده)
        // بترجع صفر لو اتقرت برمجيًا. عشان كده منقراش العمود ده خالص،
        // وبنعيد حساب الكمية بنفسنا من الأعمدة الخام (بورسعيد+
        // إسماعيلية+سويس+كارت ذكي+غاز) اللي هي أرقام عادية مش معادلات.
        final quantity = _numberValue(row, portSaidIndex) +
            _numberValue(row, ismailiaIndex) +
            _numberValue(row, suezIndex) +
            _numberValue(row, smartCardIndex) +
            _numberValue(row, gasIndex);

        final startOdo = _numberValue(row, startOdoIndex);
        final endOdo = _numberValue(row, endOdoIndex);

        final isNewCar = !carsByRecord.containsKey(recordNumber);

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

        // عداد بداية الفترة = العداد في أول شهر تظهر فيه السيارة فعليًا
        // (مش بالضرورة أول شهر في الفترة كلها)
        car.startOdometer ??= startOdo;

        // عداد نهاية الفترة = آخر عداد في آخر شهر تظهر فيه السيارة.
        // بما إننا بنمشي على الشهور بترتيبها الزمني، كل ظهور جديد
        // بيحدّث القيمة دي تلقائيًا لتبقى دايمًا "آخر" ظهور فعلي.
        car.endOdometer = endOdo;

        car.monthlyQuantities[slotIndex] = quantity;

        if (car.carNumber.isEmpty && carNumber.isNotEmpty) {
          car.carNumber = carNumber;
        }
        if (car.letters.isEmpty && letters.isNotEmpty) {
          car.letters = letters;
        }
      }
    }

    return _writeOutput(slots, carsByRecord, recordOrder);
  }

  /// بيدوّر جوه كل الشيتات المتاحة في الملف عن أول صف فيه عمود
  /// "رقم السجل" (بغض النظر عن اسم الشيت أو ترتيب الأعمدة أو رقم
  /// الصف)، ويرجّع خريطة "اسم العمود -> رقمه" عشان القراءة تبقى مرنة
  /// تمامًا مع أي شكل شيت.
  _ParsedHeader? _locateHeaderRow(ex.Excel excel) {
    for (final table in excel.tables.values) {
      final rows = table.rows;

      // العناوين غالبًا في أول 5 صفوف بس (بعد صف عنوان التقرير المدمج)
      final maxRowsToScan = rows.length < 5 ? rows.length : 5;

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

  /// بيقرا نص صف كامل ويرجّع خريطة "اسم العمود -> رقم العمود"،
  /// بغض النظر عن ترتيب الأعمدة أو مكانها في الشيت.
  Map<String, int> _findHeaders(List<ex.Data?> row) {
    final result = <String, int>{};

    for (int column = 0; column < row.length; column++) {
      final text = _cellText(row, column).trim();

      if (text.isNotEmpty) {
        result[text] = column;
      }
    }

    return result;
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

    return value.toString();
  }

  num _numberValue(List<ex.Data?> row, int? index) {
    if (index == null || index < 0 || index >= row.length) {
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

    // أي معادلة (FormulaCellValue) بترجع صفر عمدًا - مبنعتمدش على قيم
    // محسوبة جوه معادلات، خصوصًا لو الملف متفتحش في إكسل قبل القراءة.
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
  // بناء ملف الإخراج النهائي
  // =========================
  Uint8List _writeOutput(
    List<_MonthSlot> slots,
    Map<String, _CarAggregate> carsByRecord,
    List<String> recordOrder,
  ) {
    final outputExcel = ex.Excel.createExcel();
    final sheet = outputExcel['التقرير التجميعي'];

    final defaultSheet = outputExcel.getDefaultSheet();
    if (defaultSheet != null && defaultSheet != 'التقرير التجميعي') {
      outputExcel.delete(defaultSheet);
    }

    sheet.isRTL = true;

    // ===== ترتيب الأعمدة =====
    // 0: م
    // 1: رقم السجل
    // 2: رقم السيارة
    // 3: الأحرف
    // 4: عداد بداية الفترة
    // 5..(5+N-1): شهور الفترة بالترتيب
    // 5+N: عداد نهاية الفترة
    // 5+N+1: إجمالي الفترة (عداد النهاية - عداد البداية)
    // 5+N+2: إجمالي الكمية (SUM على شهور الفترة)
    final monthCount = slots.length;
    final firstMonthColumn = 4;
    final endOdoColumn = firstMonthColumn + monthCount;
    final periodTotalColumn = endOdoColumn + 1;
    final quantityTotalColumn = periodTotalColumn + 1;
    final lastColumn = quantityTotalColumn;

    final headers = <String>[
      'م',
      'رقم السجل',
      'رقم السيارة',
      'الأحرف',
      'عداد بداية الفترة',
      for (final slot in slots) slot.monthName,
      'عداد نهاية الفترة',
      'إجمالي الفترة',
      'إجمالي الكمية',
    ];

    // ===== صف العنوان =====
    final periodLabel = slots.length == 1
        ? slots.first.monthName
        : '${slots.first.monthName} إلى ${slots.last.monthName}';

    sheet.merge(
      ex.CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
      ex.CellIndex.indexByColumnRow(columnIndex: lastColumn, rowIndex: 0),
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

    // ===== صف رؤوس الأعمدة =====
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

    // ===== صفوف البيانات =====
    var outputRow = 2;
    var serial = 1;

    for (final recordNumber in recordOrder) {
      final car = carsByRecord[recordNumber]!;
      final isAltRow = serial.isOdd;

      _writeCell(sheet, 0, outputRow, serial, isAltRow: isAltRow);
      _writeCell(sheet, 1, outputRow, car.recordNumber, isAltRow: isAltRow);
      _writeCell(sheet, 2, outputRow, car.carNumber, isAltRow: isAltRow);
      _writeCell(sheet, 3, outputRow, car.letters, isAltRow: isAltRow);

      _writeCell(
        sheet,
        4,
        outputRow,
        car.startOdometer ?? 0,
        backgroundHex: _kOdometerColor,
      );

      for (int m = 0; m < monthCount; m++) {
        final quantity = car.monthlyQuantities[m] ?? 0;
        _writeCell(
          sheet,
          firstMonthColumn + m,
          outputRow,
          quantity,
          isAltRow: isAltRow,
        );
      }

      _writeCell(
        sheet,
        endOdoColumn,
        outputRow,
        car.endOdometer ?? 0,
        backgroundHex: _kOdometerColor,
      );

      final excelRow = outputRow + 1;
      final startOdoColLetter = _columnLetter(4);
      final endOdoColLetter = _columnLetter(endOdoColumn);
      final fromMonthCol = _columnLetter(firstMonthColumn);
      final toMonthCol = _columnLetter(endOdoColumn - 1);

      // إجمالي الفترة = عداد النهاية - عداد البداية (نفس الصف)
      sheet.updateCell(
        ex.CellIndex.indexByColumnRow(
          columnIndex: periodTotalColumn,
          rowIndex: outputRow,
        ),
        ex.FormulaCellValue(
          '$endOdoColLetter$excelRow-$startOdoColLetter$excelRow',
        ),
        cellStyle: _cellStyle(backgroundHex: _kTotalColor, bold: true),
      );

      // إجمالي الكمية = معادلة SUM على أعمدة الشهور بنفس الصف. الأعمدة
      // دي كلها قيم عادية كتبناها إحنا (مش معادلات مقروءة من ملف
      // خارجي)، فمفيش مشكلة caching هنا خالص.
      sheet.updateCell(
        ex.CellIndex.indexByColumnRow(
          columnIndex: quantityTotalColumn,
          rowIndex: outputRow,
        ),
        ex.FormulaCellValue(
          'SUM($fromMonthCol$excelRow:$toMonthCol$excelRow)',
        ),
        cellStyle: _cellStyle(backgroundHex: _kTotalColor, bold: true),
      );

      outputRow++;
      serial++;
    }

    // ===== عرض الأعمدة =====
    sheet.setColumnWidth(0, 6);
    sheet.setColumnWidth(1, 12);
    sheet.setColumnWidth(2, 14);
    sheet.setColumnWidth(3, 10);
    sheet.setColumnWidth(4, 14);
    for (int m = 0; m < monthCount; m++) {
      sheet.setColumnWidth(firstMonthColumn + m, 12);
    }
    sheet.setColumnWidth(endOdoColumn, 14);
    sheet.setColumnWidth(periodTotalColumn, 14);
    sheet.setColumnWidth(quantityTotalColumn, 14);

    final encoded = outputExcel.encode();

    if (encoded == null) {
      throw Exception('فشل إنشاء ملف التقرير التجميعي');
    }

    return Uint8List.fromList(encoded);
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

  ex.CellStyle _cellStyle({
    String backgroundHex = 'FFFFFFFF',
    String fontColorHex = 'FF000000',
    bool bold = false,
    int fontSize = 11,
  }) {
    final border = ex.Border(
      borderStyle: ex.BorderStyle.Thin,
      borderColorHex: ex.ExcelColor.fromHexString(_kBorderColor),
    );

    return ex.CellStyle(
      backgroundColorHex: ex.ExcelColor.fromHexString(backgroundHex),
      fontColorHex: ex.ExcelColor.fromHexString(fontColorHex),
      bold: bold,
      fontSize: fontSize,
      horizontalAlign: ex.HorizontalAlign.Center,
      verticalAlign: ex.VerticalAlign.Center,
      leftBorder: border,
      rightBorder: border,
      topBorder: border,
      bottomBorder: border,
    );
  }

  /// بيكتب القيمة الأول وبعدين التنسيق، عشان الباكدج مبيمسحش الفورمات
  /// وقت ما بيحدد صيغة الرقم تلقائيًا للخلية.
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
      ex.CellIndex.indexByColumnRow(columnIndex: column, rowIndex: row),
    );

    if (value is int) {
      cell.value = ex.IntCellValue(value);
    } else if (value is double) {
      cell.value = ex.DoubleCellValue(value);
    } else if (value is num) {
      cell.value = ex.DoubleCellValue(value.toDouble());
    } else {
      cell.value = ex.TextCellValue(value.toString());
    }

    final resolvedBackground =
        backgroundHex ?? (isAltRow ? _kAltRowColor : 'FFFFFFFF');

    cell.cellStyle = _cellStyle(
      backgroundHex: resolvedBackground,
      fontColorHex: fontColorHex,
      bold: bold,
      fontSize: fontSize,
    );
  }

  void _exportFile() {
    if (_preparedFile == null) {
      return;
    }

    String fileName = _fileNameController.text.trim();

    if (fileName.isEmpty) {
      fileName = 'التقرير التجميعي';
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
          title: const Text('التقرير التجميعي'),
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
                    Icons.summarize_outlined,
                    size: 64,
                  ),
                  const SizedBox(height: 16),
                  const Text(
                    'التقرير التجميعي',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 22,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                  const SizedBox(height: 24),

                  // اختيار الفترة
                  Row(
                    children: [
                      Expanded(
                        child: DropdownButtonFormField<int>(
                          initialValue: _fromMonth,
                          decoration: const InputDecoration(
                            labelText: 'من شهر',
                            border: OutlineInputBorder(),
                          ),
                          items: [
                            for (int m = 1; m <= 12; m++)
                              DropdownMenuItem(
                                value: m,
                                child: Text(_kArabicMonthNames[m - 1]),
                              ),
                          ],
                          onChanged: (value) {
                            setState(() {
                              _fromMonth = value;
                            });
                          },
                        ),
                      ),
                      const SizedBox(width: 12),
                      Expanded(
                        child: DropdownButtonFormField<int>(
                          initialValue: _toMonth,
                          decoration: const InputDecoration(
                            labelText: 'إلى شهر',
                            border: OutlineInputBorder(),
                          ),
                          items: [
                            for (int m = 1; m <= 12; m++)
                              DropdownMenuItem(
                                value: m,
                                child: Text(_kArabicMonthNames[m - 1]),
                              ),
                          ],
                          onChanged: (value) {
                            setState(() {
                              _toMonth = value;
                            });
                          },
                        ),
                      ),
                    ],
                  ),

                  const SizedBox(height: 12),

                  ElevatedButton(
                    onPressed: (_fromMonth == null || _toMonth == null)
                        ? null
                        : _applyPeriod,
                    child: const Text('تحديد الفترة'),
                  ),

                  const SizedBox(height: 20),

                  // خانات رفع ملف كل شهر
                  if (_monthSlots.isNotEmpty) ...[
                    const Align(
                      alignment: Alignment.centerRight,
                      child: Text(
                        'ارفع ملف كل شهر:',
                        style: TextStyle(fontWeight: FontWeight.bold),
                      ),
                    ),
                    const SizedBox(height: 8),
                    for (final slot in _monthSlots)
                      Card(
                        child: ListTile(
                          title: Text(slot.monthName),
                          subtitle: Text(
                            slot.fileName ?? 'لم يتم اختيار ملف بعد',
                            style: TextStyle(
                              color: slot.fileBytes != null
                                  ? Colors.green.shade700
                                  : Colors.grey.shade600,
                            ),
                          ),
                          trailing: IconButton(
                            icon: Icon(
                              slot.fileBytes != null
                                  ? Icons.check_circle
                                  : Icons.upload_file,
                              color: slot.fileBytes != null
                                  ? Colors.green
                                  : null,
                            ),
                            onPressed: () => _pickFileForSlot(slot),
                          ),
                          onTap: () => _pickFileForSlot(slot),
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
                    onPressed: (!_allSlotsFilled || _isProcessing)
                        ? null
                        : _generateReport,
                    icon: _isProcessing
                        ? const SizedBox(
                            width: 20,
                            height: 20,
                            child: CircularProgressIndicator(strokeWidth: 2),
                          )
                        : const Icon(Icons.summarize),
                    label: Text(
                      _isProcessing ? 'جاري التجميع...' : 'توليد التقرير',
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
