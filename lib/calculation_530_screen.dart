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
    final result = await FilePicker.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
      withData: true,
    );

    if (result == null) return;

    final file = result.files.single;

    if (file.bytes == null) {
      setState(() {
        _status = 'تعذر قراءة الملف';
      });
      return;
    }

    setState(() {
      _processing = true;
      _status = 'جاري معالجة الملف...';
    });

    try {
      final bytes = Uint8List.fromList(file.bytes!);

      final inputExcel = ex.Excel.decodeBytes(bytes);

      final sourceSheet = inputExcel.tables.containsKey('Report')
          ? inputExcel.tables['Report']!
          : inputExcel.tables[inputExcel.tables.keys.first]!;

      final outputExcel = ex.Excel.createExcel();

      final defaultSheet = outputExcel.getDefaultSheet();
      if (defaultSheet != null &&
          defaultSheet != 'سولار شهر 8' &&
          outputExcel.tables.length > 1) {
        outputExcel.delete(defaultSheet);
      }

      if (outputExcel.tables.containsKey('سولار شهر 8')) {
        outputExcel.delete('سولار شهر 8');
      }

      final sheet = outputExcel['سولار شهر 8'];

      sheet.isRTL = true;

      // =========================
      // البيانات
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

        final changedCar =
            currentGroup == null ||
            carNum != currentCarNum ||
            letters != currentLetters;

        if (changedCar) {
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
      // العنوان
      // =========================

      sheet.merge(
        ex.CellIndex.indexByString('A1'),
        ex.CellIndex.indexByString('H1'),
        customValue: ex.TextCellValue(
          'مسحوبات سولار شهر اغسطس 2026 الفترة من 1/8/2026 الي 31/8/2026',
        ),
      );

      final titleStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 16,
        bold: true,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
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

      final headerStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        bold: true,
        backgroundColorHex: ex.ExcelColor.fromHexString('D9D9D9'),
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: ex.Border(
          borderStyle: ex.BorderStyle.Thin,
        ),
        rightBorder: ex.Border(
          borderStyle: ex.BorderStyle.Thin,
        ),
        topBorder: ex.Border(
          borderStyle: ex.BorderStyle.Thin,
        ),
        bottomBorder: ex.Border(
          borderStyle: ex.BorderStyle.Thin,
        ),
      );

      for (int col = 0; col < headers.length; col++) {
        sheet.updateCell(
          ex.CellIndex.indexByColumnRow(
            columnIndex: col,
            rowIndex: 1,
          ),
          ex.TextCellValue(headers[col]),
          cellStyle: headerStyle,
        );
      }

      sheet.setRowHeight(1, 24);

      // =========================
      // البيانات
      // =========================

      int excelRow = 2;
      int serial = 1;

      final dataStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: ex.Border(
          borderStyle: ex.BorderStyle.Thin,
        ),
        rightBorder: ex.Border(
          borderStyle: ex.BorderStyle.Thin,
        ),
        topBorder: ex.Border(
          borderStyle
