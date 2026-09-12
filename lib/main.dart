import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:universal_html/html.dart' as html;
import 'excel_processor.dart';

void main() => runApp(const SmartExcelApp());

class SmartExcelApp extends StatelessWidget {
  const SmartExcelApp({super.key});
  @override
  Widget build(BuildContext context) => MaterialApp(title: 'منسق Excel الذكي', debugShowCheckedModeBanner: false, theme: ThemeData(colorScheme: ColorScheme.fromSeed(seedColor: const Color(0xff0f766e)), scaffoldBackgroundColor: const Color(0xfff4f7f5), fontFamily: 'Arial', useMaterial3: true), home: const HomePage());
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});
  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  ExcelFileInfo? reference; ExcelFileInfo? data; String status = 'في انتظار الملفات'; bool busy = false;
  Future<void> pick(bool isReference) async {
    final result = await FilePicker.pickFiles(type: FileType.custom, allowedExtensions: ['xlsx', 'xls']);
    if (result.isEmpty) return;
    final file = ExcelFileInfo(name: result.first.name, bytes: await result.first.readAsBytes());
    setState(() { if (isReference) { reference = file; status = 'تم اختيار الملف المرجعي'; } else { data = file; status = 'تم اختيار ملف البيانات الجديدة'; } });
  }
  Future<void> process() async {
    if (reference == null || data == null) { setState(() => status = 'يرجى اختيار الملفين أولًا.'); return; }
    setState(() { busy = true; status = 'جاري تحليل الملف المرجعي'; });
    try {
      final report = ExcelProcessor().process(referenceBytes: reference!.bytes, dataBytes: data!.bytes, sourceName: data!.name, onStatus: (value) => setState(() => status = value));
      final blob = html.Blob([report.bytes], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); final url = html.Url.createObjectUrlFromBlob(blob);
      html.AnchorElement(href: url)..setAttribute('download', report.fileName)..click(); html.Url.revokeObjectUrl(url); setState(() => status = 'تم الانتهاء بنجاح: ${report.summary}');
    } catch (error) { setState(() => status = 'حدث خطأ: ${error.toString().replaceFirst('FormatException: ', '')}'); } finally { setState(() => busy = false); }
  }
  @override
  Widget build(BuildContext context) => Scaffold(body: Directionality(textDirection: TextDirection.rtl, child: Center(child: ConstrainedBox(constraints: const BoxConstraints(maxWidth: 760), child: SingleChildScrollView(padding: const EdgeInsets.all(28), child: Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [
    const Text('منسق Excel الذكي', style: TextStyle(fontSize: 32, fontWeight: FontWeight.w800, color: Color(0xff123b3a))), const SizedBox(height: 8), const Text('طبّق بنية ملفك المرجعي على بيانات جديدة، محليًا وبخصوصية كاملة.', style: TextStyle(fontSize: 16, color: Color(0xff526563))), const SizedBox(height: 28),
    FileCard(title: 'الملف المرجعي', button: 'اختيار الملف المرجعي', file: reference, onPick: () => pick(true)), const SizedBox(height: 14), FileCard(title: 'ملف البيانات الجديدة', button: 'اختيار ملف البيانات الجديدة', file: data, onPick: () => pick(false)), const SizedBox(height: 22),
    FilledButton.icon(onPressed: busy ? null : process, icon: const Icon(Icons.auto_fix_high), label: const Padding(padding: EdgeInsets.symmetric(vertical: 16), child: Text('معالجة Excel', style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold)))), const SizedBox(height: 18),
    Container(padding: const EdgeInsets.all(18), decoration: BoxDecoration(color: Colors.white, border: Border.all(color: const Color(0xffd8e3df)), borderRadius: BorderRadius.circular(12)), child: Row(children: [Icon(busy ? Icons.sync : Icons.info_outline, color: const Color(0xff0f766e)), const SizedBox(width: 12), Expanded(child: Text(status, style: const TextStyle(color: Color(0xff26413f))))])), const SizedBox(height: 20), const Text('المعالجة تتم داخل المتصفح ولا يتم رفع ملفاتك إلى أي خادم.', textAlign: TextAlign.center, style: TextStyle(fontSize: 13, color: Color(0xff71817e))),
  ]))))));
}

class FileCard extends StatelessWidget {
  const FileCard({super.key, required this.title, required this.button, required this.file, required this.onPick});
  final String title; final String button; final ExcelFileInfo? file; final VoidCallback onPick;
  @override
  Widget build(BuildContext context) => Container(padding: const EdgeInsets.all(18), decoration: BoxDecoration(color: Colors.white, borderRadius: BorderRadius.circular(12), border: Border.all(color: const Color(0xffd8e3df))), child: Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [Text(title, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 17)), const SizedBox(height: 12), OutlinedButton.icon(onPressed: onPick, icon: const Icon(Icons.upload_file), label: Text(button)), if (file != null) Padding(padding: const EdgeInsets.only(top: 10), child: Text('${file!.name}  •  ${file!.sizeLabel}', style: const TextStyle(color: Color(0xff0f766e))))]));
}