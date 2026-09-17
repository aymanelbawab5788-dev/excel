import 'package:flutter/material.dart';

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
        fontFamily: 'Arial',
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color(0xff0f766e),
        ),
        scaffoldBackgroundColor: const Color(0xfff5f7f7),
      ),
      home: const HomePage(),
    );
  }
}

class HomePage extends StatelessWidget {
  const HomePage({super.key});

  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: TextDirection.rtl,
      child: Scaffold(
        appBar: AppBar(
          title: const Text(
            'منسق ملفات Excel',
            style: TextStyle(
              fontWeight: FontWeight.bold,
            ),
          ),
          centerTitle: true,
        ),
        body: Center(
          child: SingleChildScrollView(
            padding: const EdgeInsets.all(24),
            child: ConstrainedBox(
              constraints: const BoxConstraints(
                maxWidth: 900,
              ),
              child: Column(
                children: [
                  const SizedBox(height: 20),

                  const Icon(
                    Icons.table_chart_rounded,
                    size: 64,
                    color: Color(0xff0f766e),
                  ),

                  const SizedBox(height: 16),

                  const Text(
                    'منسق ملفات Excel',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 30,
                      fontWeight: FontWeight.w800,
                      color: Color(0xff173b39),
                    ),
                  ),

                  const SizedBox(height: 8),

                  const Text(
                    'اختر نوع الملف الذي تريد تجهيزه',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 16,
                      color: Color(0xff657472),
                    ),
                  ),

                  const SizedBox(height: 36),

                  LayoutBuilder(
                    builder: (context, constraints) {
                      final isWide = constraints.maxWidth >= 700;

                      if (isWide) {
                        return Row(
                          children: [
                            Expanded(
                              child: _ExcelActionCard(
                                icon: Icons.analytics_rounded,
                                title: 'كشف النسبة',
                                subtitle: 'تنسيق وتجهيز كشف النسبة',
                                onTap: () {
                                  _openSection(
                                    context,
                                    'كشف النسبة',
                                  );
                                },
                              ),
                            ),
                            const SizedBox(width: 16),
                            Expanded(
                              child: _ExcelActionCard(
                                icon: Icons.calculate_rounded,
                                title: 'حساب 530',
                                subtitle: 'حساب وتجهيز ملف 530',
                                onTap: () {
                                  _openSection(
                                    context,
                                    'حساب 530',
                                  );
                                },
                              ),
                            ),
                            const SizedBox(width: 16),
                            Expanded(
                              child: _ExcelActionCard(
                                icon: Icons.receipt_long_rounded,
                                title: 'الفواتير',
                                subtitle: 'تنسيق وتجهيز ملفات الفواتير',
                                onTap: () {
                                  _openSection(
                                    context,
                                    'الفواتير',
                                  );
                                },
                              ),
                            ),
                          ],
                        );
                      }

                      return Column(
                        children: [
                          _ExcelActionCard(
                            icon: Icons.analytics_rounded,
                            title: 'كشف النسبة',
                            subtitle: 'تنسيق وتجهيز كشف النسبة',
                            onTap: () {
                              _openSection(
                                context,
                                'كشف النسبة',
                              );
                            },
                          ),
                          const SizedBox(height: 14),
                          _ExcelActionCard(
                            icon: Icons.calculate_rounded,
                            title: 'حساب 530',
                            subtitle: 'حساب وتجهيز ملف 530',
                            onTap: () {
                              _openSection(
                                context,
                                'حساب 530',
                              );
                            },
                          ),
                          const SizedBox(height: 14),
                          _ExcelActionCard(
                            icon: Icons.receipt_long_rounded,
                            title: 'الفواتير',
                            subtitle: 'تنسيق وتجهيز ملفات الفواتير',
                            onTap: () {
                              _openSection(
                                context,
                                'الفواتير',
                              );
                            },
                          ),
                        ],
                      );
                    },
                  ),

                  const SizedBox(height: 36),

                  const Text(
                    'يتم تجهيز الملفات محليًا على جهازك',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 13,
                      color: Color(0xff7a8785),
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

  void _openSection(BuildContext context, String title) {
    Navigator.push(
      context,
      MaterialPageRoute(
        builder: (_) => TemporarySectionPage(
          title: title,
        ),
      ),
    );
  }
}

class _ExcelActionCard extends StatelessWidget {
  const _ExcelActionCard({
    required this.icon,
    required this.title,
    required this.subtitle,
    required this.onTap,
  });

  final IconData icon;
  final String title;
  final String subtitle;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return Material(
      color: Colors.white,
      borderRadius: BorderRadius.circular(18),
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(18),
        child: Container(
          padding: const EdgeInsets.all(22),
          decoration: BoxDecoration(
            borderRadius: BorderRadius.circular(18),
            border: Border.all(
              color: const Color(0xffdce5e3),
            ),
          ),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Container(
                width: 64,
                height: 64,
                decoration: BoxDecoration(
                  color: const Color(0xffe4f2ef),
                  borderRadius: BorderRadius.circular(16),
                ),
                child: Icon(
                  icon,
                  size: 34,
                  color: const Color(0xff0f766e),
                ),
              ),
              const SizedBox(height: 16),
              Text(
                title,
                textAlign: TextAlign.center,
                style: const TextStyle(
                  fontSize: 20,
                  fontWeight: FontWeight.bold,
                  color: Color(0xff173b39),
                ),
              ),
              const SizedBox(height: 6),
              Text(
                subtitle,
                textAlign: TextAlign.center,
                style: const TextStyle(
                  fontSize: 14,
                  color: Color(0xff71817e),
                ),
              ),
              const SizedBox(height: 18),
              const Icon(
                Icons.arrow_back_rounded,
                size: 22,
                color: Color(0xff0f766e),
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
        ),
        body: Center(
          child: Text(
            'صفحة $title',
            style: const TextStyle(
              fontSize: 24,
              fontWeight: FontWeight.bold,
            ),
          ),
        ),
      ),
    );
  }
}
