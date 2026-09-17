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

      if (carNumValue == null ||
          carNumValue.toString().trim().isEmpty) {
        continue;
      }

      rawRows.add({
        'car_num': _formatCarNumber(carNumValue),
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

    sheet.isRTL = true;

    // ------------------------------------------------------------
    // الألوان
    // ------------------------------------------------------------

    const black = 'FF000000';
    const darkBlue = 'FF002060';
    const red = 'FFFF0000';
    const gray = 'FFD9D9D9';
    const lightGray = 'FFF2F2F2';
    const white = 'FFFFFFFF';

    // ------------------------------------------------------------
    // الحدود
    // ------------------------------------------------------------

    final thinBorder = Border(
      borderStyle: BorderStyle.Thin,
      borderColorHex: 'FF000000',
    );

    final mediumBorder = Border(
      borderStyle: BorderStyle.Medium,
      borderColorHex: 'FF000000',
    );

    // ------------------------------------------------------------
    // الأنماط
    // ------------------------------------------------------------

    final titleStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Andalus),
      fontSize: 16,
      bold: true,
      fontColorHex: black,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
    );

    final headerStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Arial),
      fontSize: 12,
      bold: true,
      fontColorHex: black,
      backgroundColorHex: gray,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: thinBorder,
      rightBorder: thinBorder,
      topBorder: thinBorder,
      bottomBorder: thinBorder,
    );

    final dataStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Arial),
      fontSize: 11,
      bold: true,
      fontColorHex: black,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: thinBorder,
      rightBorder: thinBorder,
      topBorder: thinBorder,
      bottomBorder: thinBorder,
    );

    final blackStationStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Arial),
      fontSize: 11,
      bold: true,
      fontColorHex: black,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: thinBorder,
      rightBorder: thinBorder,
      topBorder: thinBorder,
      bottomBorder: thinBorder,
    );

    final blueStationStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Arial),
      fontSize: 11,
      bold: true,
      fontColorHex: darkBlue,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: thinBorder,
      rightBorder: thinBorder,
      topBorder: thinBorder,
      bottomBorder: thinBorder,
    );

    final sumRowStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Arial),
      fontSize: 11,
      bold: true,
      fontColorHex: black,
      backgroundColorHex: lightGray,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: thinBorder,
      rightBorder: thinBorder,
      topBorder: thinBorder,
      bottomBorder: thinBorder,
    );

    final sumValueStyle = CellStyle(
      fontFamily: getFontFamily(FontFamily.Arial),
      fontSize: 12,
      bold: true,
      fontColorHex: red,
      backgroundColorHex: lightGray,
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center,
      leftBorder: thinBorder,
      rightBorder: thinBorder,
      topBorder: thinBorder,
      bottomBorder: thinBorder,
    );

    // ------------------------------------------------------------
    // العنوان
    // ------------------------------------------------------------

    sheet.merge(
      CellIndex.indexByString('A1'),
      CellIndex.indexByString('H1'),
    );

    sheet.cell(
      CellIndex.indexByString('A1'),
    ).value = TextCellValue(
      'مسحوبات سولار شهر اغسطس 2026 الفترة من 1/8/2026 الي 31/8/2026',
    );

    sheet.cell(
      CellIndex.indexByString('A1'),
    ).cellStyle = titleStyle;

    sheet.setRowHeight(0, 30);

    // ------------------------------------------------------------
    // عناوين الأعمدة
    // ------------------------------------------------------------

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
      final cell = sheet.cell(
        CellIndex.indexByColumnRow(
          columnIndex: col,
          rowIndex: 1,
        ),
      );

      cell.value = TextCellValue(headers[col]);
      cell.cellStyle = headerStyle;
    }

    sheet.setRowHeight(1, 24);

    // ------------------------------------------------------------
    // البيانات
    // ------------------------------------------------------------

    int currentRow = 2;

    for (int groupIndex = 0;
        groupIndex < carGroups.length;
        groupIndex++) {
      final group = carGroups[groupIndex];

      final startRow = currentRow;

      for (final record in group) {
        final rowValues = [
          '',
          record['car_num'],
          record['car_letters'],
          record['station'],
          record['date'],
          record['counter'],
          record['liters'],
          record['pct'],
        ];

        final isBlackStation = [
          'بورسعيد',
          'بورفؤاد',
        ].any(
          (station) =>
              record['station'].toString().contains(station),
        );

        for (int col = 0; col < rowValues.length; col++) {
          final cell = sheet.cell(
            CellIndex.indexByColumnRow(
              columnIndex: col,
              rowIndex: currentRow,
            ),
          );

          final value = rowValues[col];

          if (value is num) {
            cell.value = DoubleCellValue(
              value.toDouble(),
            );
          } else {
            cell.value = TextCellValue(
              value?.toString() ?? '',
            );
          }

          if (col == 3) {
            cell.cellStyle = isBlackStation
                ? blackStationStyle
                : blueStationStyle;
          } else {
            cell.cellStyle = dataStyle;
          }
        }

        sheet.setRowHeight(currentRow, 18);

        currentRow++;
      }

      final endDataRow = currentRow - 1;

      // ----------------------------------------------------------
      // المسلسل المدمج
      // ----------------------------------------------------------

      final serialStart = CellIndex.indexByColumnRow(
        columnIndex: 0,
        rowIndex: startRow,
      );

      final serialEnd = CellIndex.indexByColumnRow(
        columnIndex: 0,
        rowIndex: endDataRow,
      );

      if (startRow == endDataRow) {
        sheet.cell(serialStart).value =
            IntCellValue(groupIndex + 1);

        sheet.cell(serialStart).cellStyle = dataStyle;
      } else {
        sheet.merge(
          serialStart,
          serialEnd,
        );

        sheet.cell(serialStart).value =
            IntCellValue(groupIndex + 1);

        sheet.setMergedCellStyle(
          serialStart,
          dataStyle,
        );
      }

      // ----------------------------------------------------------
      // صف المجموع
      // ----------------------------------------------------------

      for (int col = 0; col < 8; col++) {
        final cell = sheet.cell(
          CellIndex.indexByColumnRow(
            columnIndex: col,
            rowIndex: currentRow,
          ),
        );

        cell.cellStyle = sumRowStyle;
      }

      // الرقم 0 في عمود رقم السيارة
      sheet.cell(
        CellIndex.indexByColumnRow(
          columnIndex: 1,
          rowIndex: currentRow,
        ),
      ).value = IntCellValue(0);

      // معادلة مجموع اللترات
      final sumCell = sheet.cell(
        CellIndex.indexByColumnRow(
          columnIndex: 6,
          rowIndex: currentRow,
        ),
      );

      sumCell.value = FormulaCellValue(
        'SUM(G${startRow + 1}:G${endDataRow + 1})',
      );

      sumCell.cellStyle = sumValueStyle;

      sheet.setRowHeight(currentRow, 20);

      currentRow++;
    }

    // ------------------------------------------------------------
    // الإطار الخارجي العريض
    // ------------------------------------------------------------

    final maxRow = currentRow - 1;

    for (int row = 1; row <= maxRow; row++) {
      for (int col = 0; col < 8; col++) {
        final cell = sheet.cell(
          CellIndex.indexByColumnRow(
            columnIndex: col,
            rowIndex: row,
          ),
        );

        final oldStyle = cell.cellStyle;

        cell.cellStyle = oldStyle.copyWith(
          leftBorderVal:
              col == 0 ? mediumBorder : oldStyle.leftBorder,
          rightBorderVal:
              col == 7 ? mediumBorder : oldStyle.rightBorder,
          topBorderVal:
              row == 1 ? mediumBorder : oldStyle.topBorder,
          bottomBorderVal:
              row == maxRow
                  ? mediumBorder
                  : oldStyle.bottomBorder,
        );
      }
    }

    // ------------------------------------------------------------
    // عرض الأعمدة
    // ------------------------------------------------------------

    sheet.setColumnWidth(0, 7);
    sheet.setColumnWidth(1, 15);
    sheet.setColumnWidth(2, 12);
    sheet.setColumnWidth(3, 24);
    sheet.setColumnWidth(4, 16);
    sheet.setColumnWidth(5, 14);
    sheet.setColumnWidth(6, 12);
    sheet.setColumnWidth(7, 17);

    // ------------------------------------------------------------
    // حفظ الملف
    // ------------------------------------------------------------

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
    return sheet
        .cell(
          CellIndex.indexByColumnRow(
            columnIndex: column,
            rowIndex: row,
          ),
        )
        .value;
  }

  String _cellString(
    Sheet sheet,
    int row,
    int column,
  ) {
    final value = _cellValue(
      sheet,
      row,
      column,
    );

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
    final value = _cellValue(
      sheet,
      row,
      column,
    );

    if (value == null) {
      return 0;
    }

    if (value is IntCellValue) {
      return value.value.toDouble();
    }

    if (value is DoubleCellValue) {
      return value.value;
    }

    final parsed = double.tryParse(
      value.toString(),
    );

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

    html.AnchorElement(href: url)
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
