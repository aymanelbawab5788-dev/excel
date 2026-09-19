import 'package:flutter/material.dart';

import 'percentage_screen.dart';
import 'calculation_530_screen.dart';
import 'invoices_screen.dart';
import 'screens/aggregate_report_screen.dart';
import 'screens/gas_aggregate_report_screen.dart';

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

class HomePage extends StatelessWidget {
  const HomePage({super.key});

  static const List<_ExcelSection> _sections = [
    _ExcelSection(
      title: 'كشف النسبة',
      description: 'تحليل استهلاك السيارات وحساب النسب',
      icon: Icons.percent_rounded,
    ),
    _ExcelSection(
      title: 'حساب 530',
      description: 'تفريغ مسحوبات الوقود وتجهيز التقرير',
      icon: Icons.calculate_rounded,
    ),
    _ExcelSection(
      title: 'الفواتير',
      description: 'معالجة وتنسيق ملفات الفواتير',
      icon: Icons.receipt_long_rounded,
    ),
    _ExcelSection(
      title: 'التقرير التجميعي',
      description: 'تجميع بيانات السيارات عبر عدة أشهر',
      icon: Icons.summarize_outlined,
    ),
    _ExcelSection(
      title: 'التقرير التجميعي غاز',
      description: 'تجميع كميات الغاز فقط عبر عدة أشهر',
      icon: Icons.local_gas_station_outlined,
    ),
  ];

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
    return Directionality(
      textDirection: TextDirection.rtl,
      child: Scaffold(
        appBar: AppBar(
          elevation: 0,
          backgroundColor: Colors.blue,
          foregroundColor: Colors.white,
          centerTitle: true,
          title: const Text(
            'منسق ملفات Excel',
            style: TextStyle(
              fontWeight: FontWeight.bold,
              fontSize: 21,
            ),
          ),
        ),
        body: SafeArea(
          child: LayoutBuilder(
            builder: (context, constraints) {
              final width = constraints.maxWidth;

              final crossAxisCount = width < 600
                  ? 2
                  : width < 1000
                      ? 3
                      : 4;

              final horizontalPadding = width < 600 ? 16.0 : 28.0;

              return SingleChildScrollView(
                padding: EdgeInsets.fromLTRB(
                  horizontalPadding,
                  24,
                  horizontalPadding,
                  32,
                ),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    const Text(
                      'اختر العملية المطلوبة',
                      textAlign: TextAlign.center,
                      style: TextStyle(
                        fontSize: 25,
                        fontWeight: FontWeight.bold,
                      ),
                    ),
                    const SizedBox(height: 8),
                    Text(
                      'اختر أداة Excel التي تريد استخدامها',
                      textAlign: TextAlign.center,
                      style: TextStyle(
                        fontSize: 14,
                        color: Colors.grey.shade600,
                      ),
                    ),
                    const SizedBox(height: 28),
                    GridView.builder(
                      shrinkWrap: true,
                      physics: const NeverScrollableScrollPhysics(),
                      itemCount: _sections.length,
                      gridDelegate:
                          SliverGridDelegateWithFixedCrossAxisCount(
                        crossAxisCount: crossAxisCount,
                        crossAxisSpacing: 16,
                        mainAxisSpacing: 16,
                        childAspectRatio: width < 600 ? 1.05 : 1.25,
                      ),
                      itemBuilder: (context, index) {
                        final section = _sections[index];

                        return _ExcelActionCard(
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
              );
            },
          ),
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
  });

  final String title;
  final String description;
  final IconData icon;
}

class _ExcelActionCard extends StatelessWidget {
  const _ExcelActionCard({
    required this.section,
    required this.onTap,
  });

  final _ExcelSection section;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return Card(
      elevation: 1,
      shadowColor: Colors.black12,
      clipBehavior: Clip.antiAlias,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(16),
        side: BorderSide(
          color: Colors.grey.shade200,
        ),
      ),
      child: InkWell(
        onTap: onTap,
        child: Padding(
          padding: const EdgeInsets.all(18),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            mainAxisSize: MainAxisSize.min,
            children: [
              Container(
                width: 58,
                height: 58,
                decoration: BoxDecoration(
                  color: Colors.blue.withValues(alpha: 0.10),
                  borderRadius: BorderRadius.circular(16),
                ),
                child: Icon(
                  section.icon,
                  size: 30,
                  color: Colors.blue,
                ),
              ),
              const SizedBox(height: 14),
              // النص بجانب الأيقونة كان بيعمل Overflow لو العنوان طويل
              // (مثل "التقرير التجميعي غاز") لأن الـ Text ماكانش محدد له
              // maxLines، فكان بياخد سطرين ويزوّد ارتفاع الكارت عن
              // المساحة المتاحة. الحل: تحديد maxLines + Flexible
              // عشان العنوان يظهر كامل من غير ما يعمل Overflow، ومن
              // غير ما يأثر على شكل باقي الكروت اللي عنوانها قصير أصلًا.
              Flexible(
                child: Text(
                  section.title,
                  textAlign: TextAlign.center,
                  maxLines: 2,
                  overflow: TextOverflow.ellipsis,
                  style: const TextStyle(
                    fontSize: 17,
                    fontWeight: FontWeight.bold,
                    height: 1.15,
                  ),
                ),
              ),
              const SizedBox(height: 7),
              Flexible(
                child: Text(
                  section.description,
                  textAlign: TextAlign.center,
                  maxLines: 2,
                  overflow: TextOverflow.ellipsis,
                  style: TextStyle(
                    fontSize: 12,
                    height: 1.4,
                    color: Colors.grey.shade600,
                  ),
                ),
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
