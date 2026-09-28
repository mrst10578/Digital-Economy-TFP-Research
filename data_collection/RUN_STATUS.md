# وضعیت اجرای جمع‌آوری داده

آخرین نسخه‌ی اعتبارسنجی‌شده: **GitHub Actions run #23**  
تاریخ اجرا: **28 September 2026**

## نتیجه

فایل `AI_Digital_Economy_BusinessCycle_Data.xlsx` با موفقیت ساخته شد. این نسخه یک بسته‌ی داده و Audit است و **هیچ برآورد اقتصاد‌سنجی در آن اجرا نشده است**.

### خلاصه خروجی

- تعداد شیت‌ها: **132**
- وضعیت Source Log: **126 مورد OK و 0 خطا**
- ایران در `Country Summary` به‌عنوان کشور اجباری مشخص شده است.
- داده‌های ایران برای GDP، بهره‌وری نیروی کار، اشتغال، اینترنت، broadband، R&D، صادرات high-tech، تورم، بیکاری، AI Preparedness، AI patents، AI investment و DAI نگه‌داری شده‌اند.
- داده‌های BLS در اجرای بدون API Key از mirror رسمی FRED خوانده می‌شوند تا محدودیت روزانه‌ی BLS باعث افت کیفیت اجرای CI نشود.
- برای H3 علاوه بر صنایع ICT، PPI ماهانه‌ی manufacturing، mining، transportation/warehousing، wholesale و retail نیز جمع‌آوری شده است.
- طبقه‌بندی شدت دیجیتال OECD بر اساس ISIC Rev.4 داخل workbook قرار دارد.
- سهم گزارش‌شده‌ی اقتصاد دیجیتال ایران برای سال‌های 1400 و 1403 به‌صورت sparse و بدون interpolation ثبت شده است.
- نرخ‌های عملیاتی repo و corridor ایران به‌عنوان proxy مستند شده‌اند و به‌عنوان یک policy rate یکتای قطعی معرفی نشده‌اند.

## شکاف مهم باقی‌مانده

**Industrial Production ایران هنوز به‌صورت عددی بسته نشده است.**

- World Bank GEM ایران را به‌عنوان entity برمی‌گرداند، اما در استخراج فعلی برای ایران مقدار عددی IPI ندارد.
- شناسه‌های IMF Production Indexes برای ایران مستند شده‌اند:
  - `IRN.IND.IX.A` سالانه
  - `IRN.IND.IX.Q` فصلی
- این شناسه‌ها از مسیر IMF/IFS و یک mirror مشتق‌شده از IMF شناسایی شده‌اند، اما تا زمانی که یک مسیر بازتولیدپذیر برای دریافت مستقیم مقادیر عددی پیدا نشود، این شکاف **پرشده تلقی نمی‌شود**.
- شیت `Iran IPI Source Map` دقیقاً وضعیت این مسیرها را ثبت می‌کند.

## شیت‌های کنترلی اصلی

1. `Read Me`
2. `Iran Data`
3. `Iran Coverage`
4. `Iran Digital Economy`
5. `Iran IPI Source Map`
6. `Iran Monetary Rates`
7. `Country Summary`
8. `Coverage`
9. `Variable Index`
10. `Research Questions`
11. `OECD Digital Intensity`
12. `Data Dictionary`
13. `Method Notes`
14. `Source Follow-up`
15. `Source Log`

## وضعیت فرضیه‌ها

H1 تا H6 هنوز در وضعیت **partial** هستند. دلیل این وضعیت در شیت `Research Questions` برای هر فرضیه جداگانه نوشته شده است. هیچ داده‌ی گمشده‌ای با interpolation یا داده‌ی ساختگی پر نشده است.

## مرحله بعد

قبل از برآورد:

1. در صورت امکان مسیر مستقیم و قابل‌بازتولید IMF/CBI/PRC برای IPI ایران تکمیل شود.
2. crosswalk نهایی NAICS به ISIC برای H3 تصویب شود.
3. تعریف دقیق متغیرهای model و specification استاد دریافت شود.
4. سپس دیتاست تحلیلی از این workbook ساخته شود و برآورد Python آغاز شود.
