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

  Uint8List? _preparedFile;

  final TextEditingController _fileNameController =
      TextEditingController(text: 'حساب_530');

  Future<void> _pickAndProcessFile() async {
    final file = await FilePicker.pickFile(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (file == null) return;

    setState(() {
      _processing = true;
      _preparedFile = null;
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

      final sheet = outputExcel['حساب 530'];

      final defaultSheet = outputExcel.getDefaultSheet();
      if (defaultSheet != null && defaultSheet != 'حساب 530') {
        outputExcel.delete(defaultSheet);
      }

      sheet.isRTL = true;

      // =========================
      // التنسيقات
      // =========================

      final colorWhite = ex.ExcelColor.white;
      final colorTotalText = ex.ExcelColor.blue900;
      final colorTitleBg = ex.ExcelColor.blueGrey900;
      final colorHeaderBg = ex.ExcelColor.blue700;
      final colorTotalBg = ex.ExcelColor.grey300;
      final colorBorderMedium = ex.ExcelColor.grey600;
      final colorBorderThin = ex.ExcelColor.grey400;
      final colorBorderOuter = ex.ExcelColor.black;

      final thinBorder = ex.Border(
        borderStyle: ex.BorderStyle.Thin,
        borderColorHex: colorBorderThin,
      );

      final mediumBorder = ex.Border(
        borderStyle: ex.BorderStyle.Medium,
        borderColorHex: colorBorderMedium,
      );

      final outerBorder = ex.Border(
        borderStyle: ex.BorderStyle.Medium,
        borderColorHex: colorBorderOuter,
      );

      final titleStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 16,
        bold: true,
        fontColorHex: colorWhite,
        backgroundColorHex: colorTitleBg,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: mediumBorder,
        rightBorder: mediumBorder,
        topBorder: mediumBorder,
        bottomBorder: mediumBorder,
      );

      final headerStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        bold: true,
        fontColorHex: colorWhite,
        backgroundColorHex: colorHeaderBg,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        textWrapping: ex.TextWrapping.WrapText,
        leftBorder: mediumBorder,
        rightBorder: mediumBorder,
        topBorder: mediumBorder,
        bottomBorder: mediumBorder,
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

      final numericDataStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        numberFormat: ex.NumFormat.standard_2,
        leftBorder: thinBorder,
        rightBorder: thinBorder,
        topBorder: thinBorder,
        bottomBorder: thinBorder,
      );

      final totalStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        bold: true,
        fontColorHex: colorTotalText,
        backgroundColorHex: colorTotalBg,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: mediumBorder,
        rightBorder: mediumBorder,
        topBorder: mediumBorder,
        bottomBorder: mediumBorder,
      );

      final totalLabelStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        bold: true,
        italic: true,
        fontColorHex: colorTotalText,
        backgroundColorHex: colorTotalBg,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        leftBorder: mediumBorder,
        rightBorder: mediumBorder,
        topBorder: mediumBorder,
        bottomBorder: mediumBorder,
      );

      final totalNumericStyle = ex.CellStyle(
        fontFamily: ex.getFontFamily(ex.FontFamily.Arial),
        fontSize: 11,
        bold: true,
        fontColorHex: colorTotalText,
        backgroundColorHex: colorTotalBg,
        horizontalAlign: ex.HorizontalAlign.Center,
        verticalAlign: ex.VerticalAlign.Center,
        numberFormat: ex.NumFormat.standard_2,
        leftBorder: mediumBorder,
        rightBorder: mediumBorder,
        topBorder: mediumBorder,
        bottomBorder: mediumBorder,
      );

      // =========================
      // العنوان
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
            rowIndex: 0,
          ),
          ex.TextCellValue(''),
          cellStyle: titleStyle,
        );
      }

      sheet.updateCell(
        ex.CellIndex.indexByString('A1'),
        ex.TextCellValue('تفريغ مسحوبات الوقود'),
        cellStyle: titleStyle,
      );

      sheet.merge(
        ex.CellIndex.indexByString('A1'),
        ex.CellIndex.indexByString('H1'),
      );

      sheet.setMergedCellStyle(
        ex.CellIndex.indexByString('A1'),
        titleStyle,
      );

      sheet.setRowHeight(0, 26);

      // =========================
      // رؤوس الأعمدة
      // =========================

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

      sheet.setRowHeight(1, 20);

      const numericColumns = {4, 5, 6};

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
              cellStyle: numericColumns.contains(column)
                  ? numericDataStyle
                  : dataStyle,
            );
          }

          sheet.setRowHeight(rowNumber, 16);
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
            columnIndex: 3,
            rowIndex: rowNumber,
          ),
          ex.TextCellValue('الإجمالي'),
          cellStyle: totalLabelStyle,
        );

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
          cellStyle: totalNumericStyle,
        );

        sheet.setRowHeight(rowNumber, 20);

        rowNumber++;
        serial++;
      }

      // =========================
      // حد خارجي حول الجدول كله
      // =========================

      final firstBodyRow = 0;
      final lastBodyRow = rowNumber - 1;
      const firstBodyCol = 0;
      const lastBodyCol = 7;

      for (int r = firstBodyRow; r <= lastBodyRow; r++) {
        for (int c = firstBodyCol; c <= lastBodyCol; c++) {
          final onEdge = r == firstBodyRow ||
              r == lastBodyRow ||
              c == firstBodyCol ||
              c == lastBodyCol;

          if (!onEdge) continue;

          final cellIndex = ex.CellIndex.indexByColumnRow(
            columnIndex: c,
            rowIndex: r,
          );

          final existingStyle = sheet.cell(cellIndex).cellStyle;
          if (existingStyle == null) continue;

          if (r == firstBodyRow) {
            existingStyle.topBorder = outerBorder;
          }
          if (r == lastBodyRow) {
            existingStyle.bottomBorder = outerBorder;
          }
          if (c == firstBodyCol) {
            existingStyle.leftBorder = outerBorder;
          }
          if (c == lastBodyCol) {
            existingStyle.rightBorder = outerBorder;
          }

          sheet.cell(cellIndex).cellStyle = existingStyle;
        }
      }

      // =========================
      // عرض الأعمدة
      // =========================

      sheet.setColumnWidth(0, 6);
      sheet.setColumnWidth(1, 15);
      sheet.setColumnWidth(2, 12);
      sheet.setColumnWidth(3, 20);
      sheet.setColumnWidth(4, 16);
      sheet.setColumnWidth(5, 14);
      sheet.setColumnWidth(6, 13);
      sheet.setColumnWidth(7, 17);

      // =========================
      // تجهيز الملف بدون تحميل
      // =========================
      // ملاحظة مهمة: save() في مكتبة excel بتستدعي جوّاها دايمًا
      // SavingHelper.saveFile() (بتاعة الويب) حتى لو معملناش pass
      // لـ fileName — وهي دي اللي بتعمل تحميل تلقائي فورًا في المتصفح.
      // encode() بترجّع نفس البايتات بالظبط من غير ما تستدعي أي حاجة
      // خاصة بالتحميل، فمفيش تحميل إلا لما إحنا نستدعي _downloadFile()
      // بنفسنا (زر "تصدير الملف").
      final outputBytes = outputExcel.encode();

      if (outputBytes == null) {
        throw Exception('فشل إنشاء ملف Excel');
      }

      setState(() {
        _preparedFile = Uint8List.fromList(outputBytes);
        _processing = false;
        _status = 'تم تجهيز الملف بنجاح — اضغط تصدير الملف';
      });
    } catch (e) {
      setState(() {
        _processing = false;
        _preparedFile = null;
        _status = 'حدث خطأ: $e';
      });
    }
  }

  void _exportFile() {
    if (_preparedFile == null) return;

    String fileName = _fileNameController.text.trim();

    if (fileName.isEmpty) {
      fileName = 'حساب_530';
    }

    if (!fileName.toLowerCase().endsWith('.xlsx')) {
      fileName = '$fileName.xlsx';
    }

    _downloadFile(
      _preparedFile!,
      fileName,
    );

    setState(() {
      _status = 'تم تصدير الملف بنجاح';
    });
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
  void dispose() {
    _fileNameController.dispose();
    super.dispose();
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
          child: SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: ConstrainedBox(
              constraints: const BoxConstraints(
                maxWidth: 550,
              ),
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

                  // اختيار الملف
                  SizedBox(
                    width: double.infinity,
                    child: ElevatedButton.icon(
                      onPressed:
                          _processing ? null : _pickAndProcessFile,
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
                  ),

                  const SizedBox(height: 20),

                  // اسم الملف
                  TextField(
                    controller: _fileNameController,
                    textDirection: TextDirection.rtl,
                    enabled: _preparedFile != null && !_processing,
                    decoration: const InputDecoration(
                      labelText: 'اسم الملف',
                      hintText: 'اكتب اسم الملف',
                      prefixIcon: Icon(
                        Icons.drive_file_rename_outline,
                      ),
                      border: OutlineInputBorder(),
                    ),
                  ),

                  const SizedBox(height: 12),

                  Text(
                    _status,
                    textAlign: TextAlign.center,
                  ),

                  const SizedBox(height: 20),

                  // تصدير الملف
                  SizedBox(
                    width: double.infinity,
                    child: FilledButton.icon(
                      onPressed:
                          _preparedFile == null || _processing
                              ? null
                              : _exportFile,
                      icon: const Icon(Icons.download),
                      label: const Text('تصدير الملف'),
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
