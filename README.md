# Qarenly Full Stack Starter

منصة Next.js كاملة كبداية عملية لمشروع مقارنة أسعار وخصومات الأصناف.

## الموجود داخل المشروع
- تسجيل دخول وإنشاء حساب مورد
- حساب أدمن تجريبي
- API routes للـ login / register / offers / compare
- رفع عروض عبر Excel
- مقارنة ملفين Excel
- محرك بحث سوقي
- لوحة مورد
- لوحة أدمن

## بيانات الدخول التجريبية
- admin@qarenly.com
- 123456

## التشغيل
```bash
npm install
npm run dev
```

ثم افتح:
```bash
http://localhost:3000
```

## ملاحظات مهمة
- التخزين الحالي In-Memory كبداية سريعة للعرض والتطوير.
- الخطوة التالية الطبيعية: ربط PostgreSQL + Prisma + JWT/Cookies + S3/Blob storage.
- ملفات Excel المطلوبة تحتوي على الأعمدة:
  - name
  - price
  - discount

## ما الذي تطوره بعد ذلك
1. قاعدة بيانات حقيقية
2. رفع ملفات محفوظة
3. صلاحيات أدق
4. اشتراكات ودفع
5. تنبيهات وتقارير
