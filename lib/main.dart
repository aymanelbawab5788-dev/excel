import 'package:flutter/material.dart';

import 'percentage_screen.dart';
import 'calculation_530_screen.dart';
import 'invoices_screen.dart';
import 'screens/aggregate_report_screen.dart';
import 'screens/gas_aggregate_report_screen.dart';
import 'screens/add_vehicle_data_screen.dart';

void main() {
  runApp(const ExcelFormatterApp());
}

class ExcelFormatterApp extends StatelessWidget {
  const ExcelFormatterApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'منسق ملفات Excel',
      theme: ThemeData(
        useMaterial3: true,
        colorSchemeSeed: Colors.blue,
        scaffoldBackgroundColor: const Color(0xFFF6F8FB),
      ),
      home: const HomePage(),
    );
  }
}

// ============================================================
// تصميم الصفحة الرئيسية من جديد:
//
// السبب إننا سبنا شكل "الكروت المربعة" (Grid) اللي كان بيعمل
// Overflow، وبدّلناه بقائمة صفوف (List) بارتفاع مرن (مش ثابت).
// ده حل جذري لمشكلة الأوفر فلو نفسها: كل صف بياخد بالظبط
// الارتفاع اللي محتاجه حسب طول النص وحجم خط الجهاز، فمفيش أي
// احتمال إن النص يتقطع أو يطلع بره حدود الصف تاني، مهما كان
// عنوان القسم طويل ومهما كان إعداد حجم الخط في الموبايل.
//
// وبما إننا لسه هنضيف أقسام جديدة، القائمة دي هتستوعب أي عدد
// أقسام من غير ما نحتاج نعدّل أي حسابات (زي عدد الأعمدة أو
// الـ aspect ratio اللي كانت في التصميم القديم)، وفيها خانة
// بحث بسيطة تساعد لما عدد الأدوات يكبر.
// ============================================================

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  final TextEditingController _searchController = TextEditingController();
  String _query = '';

  static const List<_ExcelSection> _sections = [
    _ExcelSection(
      title: 'كشف النسبة',
      description: 'تحليل استهلاك السيارات وحساب النسب',
      icon: Icons.percent_rounded,
      color: Color(0xFF2E7D32),
    ),
    _ExcelSection(
      title: 'حساب 530',
      description: 'تفريغ مسحوبات الوقود وتجهيز التقرير',
      icon: Icons.calculate_rounded,
      color: Color(0xFF6A1B9A),
    ),
    _ExcelSection(
      title: 'الفواتير',
      description: 'معالجة وتنسيق ملفات الفواتير',
      icon: Icons.receipt_long_rounded,
      color: Color(0xFFEF6C00),
    ),
    _ExcelSection(
      title: 'التقرير التجميعي',
      description: 'تجميع بيانات السيارات عبر عدة أشهر',
      icon: Icons.summarize_outlined,
      color: Color(0xFF1565C0),
    ),
    _ExcelSection(
      title: 'التقرير التجميعي غاز',
      description: 'تجميع كميات الغاز فقط عبر عدة أشهر',
      icon: Icons.local_gas_station_outlined,
      color: Color(0xFFC62828),
    ),
    _ExcelSection(
      title: 'إضافة بيانات للتقارير',
      description: 'إضافة بيانات المركبات كأعمدة جديدة لأي تقرير',
      icon: Icons.playlist_add_rounded,
      color: Color(0xFF00897B),
    ),
  ];

  List<_ExcelSection> get _filteredSections {
    final query = _query.trim();

    if (query.isEmpty) {
      return _sections;
    }

    return _sections
        .where(
          (section) =>
              section.title.contains(query) ||
              section.description.contains(query),
        )
        .toList();
  }

  @override
  void dispose() {
    _searchController.dispose();
    super.dispose();
  }

  void _openSection(BuildContext context, String title) {
    if (title == 'كشف النسبة') {
      Navigator.push(
        context,
        MaterialPageRoute(
          builder: (_) => const PercentageScreen(),
        ),
      );
      return;
    }

    if (title == 'حساب 530') {
      Navigator.push(
        context,
        MaterialPageRoute(
          builder: (_) => const Calculation530Screen(),
        ),
      );
      return;
    }

    if (title == 'الفواتير') {
      Navigator.push(
        context,
        MaterialPageRoute(
          builder: (_) => const InvoicesScreen(),
        ),
      );
      return;
    }

    if (title == 'التقرير التجميعي') {
      Navigator.push(
        context,
        MaterialPageRoute(
          builder: (_) => const AggregateReportScreen(),
        ),
      );
      return;
    }

    if (title == 'التقرير التجميعي غاز') {
      Navigator.push(
        context,
        MaterialPageRoute(
          builder: (_) => const GasAggregateReportScreen(),
        ),
      );
      return;
    }

    if (title == 'إضافة بيانات للتقارير') {
      Navigator.push(
        context,
        MaterialPageRoute(
          builder: (_) => const AddVehicleDataScreen(),
        ),
      );
      return;
    }

    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (_) => TemporarySectionPage(
          title: title,
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final filtered = _filteredSections;

    return Directionality(
      textDirection: TextDirection.rtl,
      child: Scaffold(
        body: CustomScrollView(
          slivers: [
            // =========================
            // الهيدر
            // =========================
            SliverAppBar(
              pinned: true,
              stretch: true,
              expandedHeight: 150,
              backgroundColor: Colors.blue,
              foregroundColor: Colors.white,
              elevation: 0,
              flexibleSpace: FlexibleSpaceBar(
                titlePadding: const EdgeInsetsDirectional.only(
                  start: 20,
                  bottom: 16,
                ),
                title: const Text(
                  'منسق ملفات Excel',
                  style: TextStyle(
                    fontWeight: FontWeight.bold,
                    fontSize: 20,
                  ),
                ),
                background: Container(
                  decoration: const BoxDecoration(
                    gradient: LinearGradient(
                      begin: Alignment.topRight,
                      end: Alignment.bottomLeft,
                      colors: [
                        Color(0xFF1565C0),
                        Color(0xFF1E88E5),
                      ],
                    ),
                  ),
                  child: Align(
                    alignment: AlignmentDirectional.bottomStart,
                    child: Padding(
                      padding: const EdgeInsetsDirectional.only(
                        start: 20,
                        bottom: 46,
                      ),
                      child: Icon(
                        Icons.grid_view_rounded,
                        size: 34,
                        color: Colors.white.withValues(alpha: 0.35),
                      ),
                    ),
                  ),
                ),
              ),
            ),

            // =========================
            // محتوى الصفحة
            // =========================
            SliverToBoxAdapter(
              child: Center(
                child: ConstrainedBox(
                  constraints: const BoxConstraints(maxWidth: 680),
                  child: Padding(
                    padding: const EdgeInsets.fromLTRB(20, 22, 20, 32),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      children: [
                        const Text(
                          'اختر العملية المطلوبة',
                          textAlign: TextAlign.center,
                          style: TextStyle(
                            fontSize: 23,
                            fontWeight: FontWeight.bold,
                          ),
                        ),
                        const SizedBox(height: 6),
                        Text(
                          'اختر أداة Excel التي تريد استخدامها',
                          textAlign: TextAlign.center,
                          style: TextStyle(
                            fontSize: 13.5,
                            color: Colors.grey.shade600,
                          ),
                        ),
                        const SizedBox(height: 22),

                        // خانة البحث — مفيدة أكتر كل ما عدد الأدوات يزيد
                        TextField(
                          controller: _searchController,
                          textDirection: TextDirection.rtl,
                          onChanged: (value) {
                            setState(() {
                              _query = value;
                            });
                          },
                          decoration: InputDecoration(
                            hintText: 'ابحث عن أداة...',
                            prefixIcon: const Icon(Icons.search_rounded),
                            suffixIcon: _query.isEmpty
                                ? null
                                : IconButton(
                                    icon: const Icon(Icons.close_rounded),
                                    onPressed: () {
                                      setState(() {
                                        _searchController.clear();
                                        _query = '';
                                      });
                                    },
                                  ),
                            filled: true,
                            fillColor: Colors.white,
                            contentPadding: const EdgeInsets.symmetric(
                              vertical: 14,
                            ),
                            border: OutlineInputBorder(
                              borderRadius: BorderRadius.circular(14),
                              borderSide: BorderSide(
                                color: Colors.grey.shade300,
                              ),
                            ),
                            enabledBorder: OutlineInputBorder(
                              borderRadius: BorderRadius.circular(14),
                              borderSide: BorderSide(
                                color: Colors.grey.shade300,
                              ),
                            ),
                          ),
                        ),

                        const SizedBox(height: 20),

                        // =========================
                        // قائمة الأدوات
                        // =========================
                        if (filtered.isEmpty)
                          Padding(
                            padding: const EdgeInsets.symmetric(
                              vertical: 40,
                            ),
                            child: Column(
                              children: [
                                Icon(
                                  Icons.search_off_rounded,
                                  size: 40,
                                  color: Colors.grey.shade400,
                                ),
                                const SizedBox(height: 10),
                                Text(
                                  'مفيش أدوات مطابقة للبحث',
                                  style: TextStyle(
                                    color: Colors.grey.shade600,
                                  ),
                                ),
                              ],
                            ),
                          )
                        else
                          ListView.separated(
                            shrinkWrap: true,
                            physics: const NeverScrollableScrollPhysics(),
                            itemCount: filtered.length,
                            separatorBuilder: (_, _) =>
                                const SizedBox(height: 12),
                            itemBuilder: (context, index) {
                              final section = filtered[index];

                              return _ExcelActionTile(
                                section: section,
                                onTap: () => _openSection(
                                  context,
                                  section.title,
                                ),
                              );
                            },
                          ),
                      ],
                    ),
                  ),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}

class _ExcelSection {
  const _ExcelSection({
    required this.title,
    required this.description,
    required this.icon,
    required this.color,
  });

  final String title;
  final String description;
  final IconData icon;
  final Color color;
}

/// صف عملية واحدة — ارتفاعه مرن حسب المحتوى، مش ثابت، عشان
/// النص (مهما طال، ومهما كبر حجم خط الجهاز) يفضل ظاهر كامل
/// من غير أي Overflow.
class _ExcelActionTile extends StatelessWidget {
  const _ExcelActionTile({
    required this.section,
    required this.onTap,
  });

  final _ExcelSection section;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return Material(
      color: Colors.white,
      borderRadius: BorderRadius.circular(18),
      child: InkWell(
        borderRadius: BorderRadius.circular(18),
        onTap: onTap,
        child: Container(
          padding: const EdgeInsets.all(14),
          decoration: BoxDecoration(
            borderRadius: BorderRadius.circular(18),
            border: Border.all(color: Colors.grey.shade200),
          ),
          child: Row(
            crossAxisAlignment: CrossAxisAlignment.center,
            children: [
              Container(
                width: 52,
                height: 52,
                decoration: BoxDecoration(
                  color: section.color.withValues(alpha: 0.12),
                  borderRadius: BorderRadius.circular(14),
                ),
                child: Icon(
                  section.icon,
                  size: 26,
                  color: section.color,
                ),
              ),
              const SizedBox(width: 14),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    Text(
                      section.title,
                      textAlign: TextAlign.right,
                      style: const TextStyle(
                        fontSize: 16,
                        fontWeight: FontWeight.bold,
                        height: 1.25,
                      ),
                    ),
                    const SizedBox(height: 3),
                    Text(
                      section.description,
                      textAlign: TextAlign.right,
                      style: TextStyle(
                        fontSize: 12.5,
                        height: 1.35,
                        color: Colors.grey.shade600,
                      ),
                    ),
                  ],
                ),
              ),
              const SizedBox(width: 6),
              Icon(
                Icons.chevron_left_rounded,
                color: Colors.grey.shade400,
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class TemporarySectionPage extends StatelessWidget {
  const TemporarySectionPage({
    super.key,
    required this.title,
  });

  final String title;

  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: TextDirection.rtl,
      child: Scaffold(
        appBar: AppBar(
          title: Text(title),
          centerTitle: true,
        ),
        body: Center(
          child: Text(
            'قسم $title سيتم تنفيذه لاحقًا',
            style: const TextStyle(
              fontSize: 20,
            ),
          ),
        ),
      ),
    );
  }
}
