import 'dart:typed_data';

import 'package:excel/excel.dart' as ex;
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
  String _status = 'اختر ملف Excel للبدء';

  Future<void> _pickAndProcessFile() async {
    final file = await FilePicker.pickFile(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (file == null) return;

    setState(() {
      _processing = true;
      _status = 'جاري معالجة الملف...';
    });

    try {
      final bytes = await file.readAsBytes();
      final inputExcel = ex.Excel.decodeBytes(bytes);

      if (inputExcel.tables.isEmpty) {
        throw Exception('الملف لا يحتوي على أي Sheet');
      }

      final sourceSheet = inputExcel.tables.containsKey('Report')
          ? inputExcel.tables['Report']!
          : inputExcel.tables.values.first;

      final outputExcel = ex.Excel.createExcel();

      final defaultSheet = outputExcel.getDefaultSheet();
      if (defaultSheet != null) {
        outputExcel.delete(defaultSheet);
      }

      final sheet = outputExcel['حساب 530'];

      // =========================
      // التنسيقات البسيطة
      // =========================

      final thinBorder = ex.Border(
        borderStyle: ex.BorderStyle.Thin,
      );

      final titleStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 16,
        bold: true,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
      );

      final headerStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        bold: true,
        backgroundColorHex: ex.ExcelColor.fromHexString('D9D9D9'),
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: thinBorder,
        rightBorder: thinBorder,
        topBorder: thinBorder,
        bottomBorder: thinBorder,
      );

      final dataStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: thinBorder,
        rightBorder: thinBorder,
        topBorder: thinBorder,
        bottomBorder: thinBorder,
      );

      final totalStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 12,
        bold: true,
        fontColorHex: ex.ExcelColor.fromHexString('FF0000'),
        backgroundColorHex: ex.ExcelColor.fromHexString('F2F2F2'),
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: thinBorder,
        rightBorder: thinBorder,
        topBorder: thinBorder,
        bottomBorder: thinBorder,
      );

      // =========================
      // العنوان
      // =========================

      sheet.merge(
        ex.CellIndex.indexByString('A1'),
        ex.CellIndex.indexByString('H1'),
      );

      sheet.updateCell(
        ex.CellIndex.indexByString('A1'),
        ex.TextCellValue('تفريغ مسحوبات الوقود'),
        cellStyle: titleStyle,
      );

      sheet.setMergedCellStyle(
        ex.CellIndex.indexByString('A1'),
        titleStyle,
      );

      sheet.setRowHeight(0, 30);

      // =========================
      // رؤوس الأعمدة
      // =========================

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

      for (int column = 0; column < headers.length; column++) {
        sheet.updateCell(
          ex.CellIndex.indexByColumnRow(
            columnIndex: column,
            rowIndex: 1,
          ),
          ex.TextCellValue(headers[column]),
          cellStyle: headerStyle,
        );
      }

      sheet.setRowHeight(1, 24);

      // =========================
      // تجميع السيارات
      // =========================

      final groups = <List<List<dynamic>>>[];

      List<List<dynamic>>? currentGroup;
      String? currentCarNum;
      String? currentLetters;

      for (int rowIndex = 1; rowIndex < sourceSheet.maxRows; rowIndex++) {
        final row = sourceSheet.row(rowIndex);

        if (row.length < 7) continue;

        final carNum = _cellText(row, 0);
        final letters = _cellText(row, 1);

        if (carNum.isEmpty && letters.isEmpty) {
          continue;
        }

        final station = _cellText(row, 2);
        final date = _cellText(row, 3);
        final counter = _cellValue(row, 4);
        final liters = _cellValue(row, 5);
        final percentage = _cellValue(row, 6);

        final newCar =
            currentGroup == null ||
            carNum != currentCarNum ||
            letters != currentLetters;

        if (newCar) {
          currentGroup = [];
          groups.add(currentGroup);

          currentCarNum = carNum;
          currentLetters = letters;
        }

        currentGroup.add([
          carNum,
          letters,
          station,
          date,
          counter,
          liters,
          percentage,
        ]);
      }

      // =========================
      // كتابة البيانات
      // =========================

      int rowNumber = 2;
      int serial = 1;

      for (final group in groups) {
        final firstDataRow = rowNumber;

        for (final item in group) {
          final values = [
            serial,
            item[0],
            item[1],
            item[2],
            item[3],
            item[4],
            item[5],
            item[6],
          ];

          for (int column = 0; column < values.length; column++) {
            final value = values[column];

            ex.CellValue cellValue;

            if (value is int) {
              cellValue = ex.IntCellValue(value);
            } else if (value is double) {
              cellValue = ex.DoubleCellValue(value);
            } else if (value is num) {
              cellValue = ex.DoubleCellValue(value.toDouble());
            } else {
              cellValue = ex.TextCellValue(value.toString());
            }

            sheet.updateCell(
              ex.CellIndex.indexByColumnRow(
                columnIndex: column,
                rowIndex: rowNumber,
              ),
              cellValue,
              cellStyle: dataStyle,
            );
          }

          sheet.setRowHeight(rowNumber, 18);
          rowNumber++;
        }

        final lastDataRow = rowNumber - 1;

        // دمج رقم المسلسل
        if (lastDataRow > firstDataRow) {
          sheet.merge(
            ex.CellIndex.indexByColumnRow(
              columnIndex: 0,
              rowIndex: firstDataRow,
            ),
            ex.CellIndex.indexByColumnRow(
              columnIndex: 0,
              rowIndex: lastDataRow,
            ),
          );

          sheet.setMergedCellStyle(
            ex.CellIndex.indexByColumnRow(
              columnIndex: 0,
              rowIndex: firstDataRow,
            ),
            dataStyle,
          );
        }

        // صف الإجمالي
        for (int column = 0; column < 8; column++) {
          sheet.updateCell(
            ex.CellIndex.indexByColumnRow(
              columnIndex: column,
              rowIndex: rowNumber,
            ),
            ex.TextCellValue(''),
            cellStyle: totalStyle,
          );
        }

        sheet.updateCell(
          ex.CellIndex.indexByColumnRow(
            columnIndex: 1,
            rowIndex: rowNumber,
          ),
          ex.IntCellValue(0),
          cellStyle: totalStyle,
        );

        sheet.updateCell(
          ex.CellIndex.indexByColumnRow(
            columnIndex: 6,
            rowIndex: rowNumber,
          ),
          ex.FormulaCellValue(
            'SUM(G${firstDataRow + 1}:G${lastDataRow + 1})',
          ),
          cellStyle: totalStyle,
        );

        sheet.setRowHeight(rowNumber, 26);

        rowNumber++;
        serial++;
      }

      // =========================
      // عرض الأعمدة
      // =========================

      sheet.setColumnWidth(0, 7);
      sheet.setColumnWidth(1, 14);
      sheet.setColumnWidth(2, 12);
      sheet.setColumnWidth(3, 18);
      sheet.setColumnWidth(4, 15);
      sheet.setColumnWidth(5, 14);
      sheet.setColumnWidth(6, 12);
      sheet.setColumnWidth(7, 16);

      // =========================
      // تحميل الملف
      // =========================

      final outputBytes = outputExcel.save(
        fileName: 'حساب_530.xlsx',
      );

      if (outputBytes == null) {
        throw Exception('فشل إنشاء ملف Excel');
      }

      _downloadFile(
        Uint8List.fromList(outputBytes),
        'حساب_530.xlsx',
      );

      setState(() {
        _processing = false;
        _status = 'تم تجهيز الملف وتحميله بنجاح';
      });
    } catch (e) {
      setState(() {
        _processing = false;
        _status = 'حدث خطأ: $e';
      });
    }
  }

  String _cellText(List<ex.Data?> row, int index) {
    if (index >= row.length) return '';

    final data = row[index];

    if (data == null || data.value == null) {
      return '';
    }

    return data.value.toString();
  }

  dynamic _cellValue(List<ex.Data?> row, int index) {
    if (index >= row.length) return '';

    final data = row[index];

    if (data == null || data.value == null) {
      return '';
    }

    final value = data.value;

    if (value is ex.IntCellValue) {
      return value.value;
    }

    if (value is ex.DoubleCellValue) {
      return value.value;
    }

    return value.toString();
  }

  void _downloadFile(Uint8List bytes, String fileName) {
    final blob = html.Blob([bytes]);
    final url = html.Url.createObjectUrlFromBlob(blob);

    html.AnchorElement(href: url)
      ..setAttribute('download', fileName)
      ..click();

    html.Url.revokeObjectUrl(url);
  }

  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: TextDirection.rtl,
      child: Scaffold(
        appBar: AppBar(
          title: const Text('حساب 530'),
          centerTitle: true,
        ),
        body: Center(
          child: Padding(
            padding: const EdgeInsets.all(24),
            child: Column(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                const Icon(
                  Icons.calculate_outlined,
                  size: 70,
                ),
                const SizedBox(height: 20),
                const Text(
                  'حساب 530',
                  style: TextStyle(
                    fontSize: 24,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                const SizedBox(height: 30),
                ElevatedButton.icon(
                  onPressed: _processing ? null : _pickAndProcessFile,
                  icon: _processing
                      ? const SizedBox(
                          width: 20,
                          height: 20,
                          child: CircularProgressIndicator(
                            strokeWidth: 2,
                          ),
                        )
                      : const Icon(Icons.upload_file),
                  label: Text(
                    _processing
                        ? 'جاري المعالجة...'
                        : 'اختيار ملف Excel',
                  ),
                ),
                const SizedBox(height: 20),
                Text(
                  _status,
                  textAlign: TextAlign.center,
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
