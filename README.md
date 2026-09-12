# منسق Excel الذكي

تطبيق Flutter Web لمعالجة ملفات Excel محليًا داخل المتصفح. يقرأ ملفًا مرجعيًا وملف بيانات جديدًا، يكتشف بنية الجداول والأعمدة، ثم ينشئ نسخة منسقة قابلة للتنزيل.

## التشغيل

```bash
flutter pub get
flutter run -d chrome
```

## البناء

```bash
flutter build web --release
```

المعالجة لا تستخدم Backend أو API أو قاعدة بيانات، ولا تغادر الملفات جهاز المستخدم. ناتج البناء موجود في `build/web` ويمكن رفعه كملفات Static إلى Hugging Face Spaces.