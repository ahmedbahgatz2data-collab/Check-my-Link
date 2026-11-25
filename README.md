# URL Checker Pro — By ENG Ahmed Bahgat

أداة سطح مكتب احترافية بلغة Python لفحص الروابط وتتبع التحويلات النهائية (HTTP Redirects & JavaScript/Meta Redirects) باستخدام `requests` و `Playwright`.

## المميزات
- **وضع الفحص الذكي (`Smart HTTP + Playwright JS`):** فحص سريع عبر HTTP مع التحويل التلقائي إلى متصفح Chromium عبر Playwright عند الحاجة.
- **فحص متوازي (Multi-Threading):** دعم تشغيل عدة خيوط في نفس الوقت لتسريع الفحص.
- **تتبع كامل للتحويلات:** استخراج الرابط النهائي، عدد القفزات، سلسلة التحويلات كاملة، عنوان الصفحة، حجم المحتوى، وزمن الاستجابة.
- **تصدير متعدد الصيغ:** `.xlsx` و `.csv` و `.txt` (Tab-delimited).
- **بناء تلقائي عبر GitHub Actions:** يقوم ملف `.github/workflows/build.yml` بتجميع البرنامج تلقائياً إلى `.exe` ونشره في قسم **Releases** مع كل رفع إلى `main` أو عند إنشاء Tag جديد.
