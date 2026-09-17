import 'package:flutter/material.dart';

import 'percentage_screen.dart';

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
      ),
      home: const HomePage(),
    );
  }
}

class HomePage extends StatelessWidget {
  const HomePage({super.key});

  void _openSection(
    BuildContext context,
    String title,
  ) {
    if (title == 'كشف النسبة') {
      Navigator.push(
        context,
        MaterialPageRoute(
          builder: (_) => const PercentageScreen(),
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
          title: const Text('منسق ملفات Excel'),
          centerTitle: true,
        ),
        body: SafeArea(
          child: Padding(
            padding: const EdgeInsets.all(20),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                const SizedBox(height: 20),

                const Text(
                  'اختر العملية المطلوبة',
                  textAlign: TextAlign.center,
                  style: TextStyle(
                    fontSize: 24,
                    fontWeight: FontWeight.bold,
                  ),
                ),

                const SizedBox(height: 30),

                Expanded(
                  child: LayoutBuilder(
                    builder: (context, constraints) {
                      final isWide = constraints.maxWidth >= 700;

                      if (isWide) {
                        return Row(
                          crossAxisAlignment: CrossAxisAlignment.stretch,
                          children: [
                            Expanded(
                              child: _ExcelActionCard(
                                title: 'كشف النسبة',
                                icon: Icons.percent,
                                onTap: () => _openSection(
                                  context,
                                  'كشف النسبة',
                                ),
                              ),
                            ),
                            const SizedBox(width: 16),
                            Expanded(
                              child: _ExcelActionCard(
                                title: 'حساب 530',
                                icon: Icons.calculate_outlined,
                                onTap: () => _openSection(
                                  context,
                                  'حساب 530',
                                ),
                              ),
                            ),
                            const SizedBox(width: 16),
                            Expanded(
                              child: _ExcelActionCard(
                                title: 'الفواتير',
                                icon: Icons.receipt_long_outlined,
                                onTap: () => _openSection(
                                  context,
                                  'الفواتير',
                                ),
                              ),
                            ),
                          ],
                        );
                      }

                      return ListView(
                        children: [
                          _ExcelActionCard(
                            title: 'كشف النسبة',
                            icon: Icons.percent,
                            onTap: () => _openSection(
                              context,
                              'كشف النسبة',
                            ),
                          ),
                          const SizedBox(height: 16),
                          _ExcelActionCard(
                            title: 'حساب 530',
                            icon: Icons.calculate_outlined,
                            onTap: () => _openSection(
                              context,
                              'حساب 530',
                            ),
                          ),
                          const SizedBox(height: 16),
                          _ExcelActionCard(
                            title: 'الفواتير',
                            icon: Icons.receipt_long_outlined,
                            onTap: () => _openSection(
                              context,
                              'الفواتير',
                            ),
                          ),
                        ],
                      );
                    },
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}

class _ExcelActionCard extends StatelessWidget {
  const _ExcelActionCard({
    required this.title,
    required this.icon,
    required this.onTap,
  });

  final String title;
  final IconData icon;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return Card(
      elevation: 2,
      child: InkWell(
        onTap: onTap,
        borderRadius: BorderRadius.circular(12),
        child: Padding(
          padding: const EdgeInsets.all(24),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              Icon(
                icon,
                size: 50,
              ),
              const SizedBox(height: 16),
              Text(
                title,
                textAlign: TextAlign.center,
                style: const TextStyle(
                  fontSize: 20,
                  fontWeight: FontWeight.bold,
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
    return Scaffold(
      appBar: AppBar(
        title: Text(title),
      ),
      body: Center(
        child: Text(
          'قسم $title سيتم تنفيذه لاحقًا',
          style: const TextStyle(
            fontSize: 20,
          ),
        ),
      ),
    );
  }
}

بعد وضعه، لازم يكون عندنا:

lib/
├── main.dart
└── percentage_screen.dart

وبكده الضغط على كشف النسبة هيفتح "PercentageScreen"، أما حساب 530 والفواتير فمؤقتين زي ما اتفقنا.

بعدها نجرب التشغيل، ولو اشتغل نبدأ في منطق كشف النسبة.
