# AI, Digital Economy & Business Cycle Dynamics

این ریپو برای جمع‌آوری، مستندسازی و آماده‌سازی داده‌های پروژه‌ی اقتصاد دیجیتال، هوش مصنوعی و پویایی چرخه‌های تجاری ساخته شده است.

## فاز فعلی: Data Collection + Coverage Audit

در این فاز **مدل اقتصاد‌سنجی نهایی اجرا نمی‌شود**. ابتدا داده‌های موردنیاز از منابع عمومی جمع‌آوری می‌شوند، پوشش کشور/زمان/فرکانس آن‌ها Audit می‌شود و شکاف‌های واقعی بدون ساختن داده‌ی مصنوعی ثبت می‌شوند.

اسکریپت اصلی:

`data_collection/collect_business_cycle_data.py`

خروجی اصلی:

`AI_Digital_Economy_BusinessCycle_Data.xlsx`

## منابع فعال

- FRED
- World Bank WDI
- World Bank Global Economic Monitor
- BEA public downloads / API
- BLS series، با fallback بازتولیدپذیر از mirror رسمی FRED در محیط بدون API Key
- OECD
- IMF DataMapper
- World Bank Digital Adoption Index
- Census Business Trends and Outlook Survey
- CSET / ETO Country AI Activity Metrics
- Epoch AI
- AIOE / AIIE exposure datasets

## ایران

ایران عمداً در Audit کشورها اجباری است و شیت‌های اختصاصی زیر دارد:

- `Iran Data`
- `Iran Coverage`
- `Iran Digital Economy`
- `Iran IPI Source Map`
- `Iran Monetary Rates`

در خروجی فعلی، داده‌های رسمی/منبع‌دار برای بخش بزرگی از متغیرهای کلیدی ایران موجود است. Industrial Production ایران هنوز یک شکاف واقعی است و به‌جای پر کردن مصنوعی، مسیرهای IMF PI، World Bank GEM و منبع داخلی در شیت `Iran IPI Source Map` مستند شده‌اند.

## H1 تا H6

شیت `Research Questions` نشان می‌دهد برای هر فرضیه چه داده‌ای موجود است و چه چیزی باقی مانده. شیت `Variable Index` نیز وضعیت هر متغیر را به‌صورت available / partial / missing ثبت می‌کند.

## اجرای محلی

```bash
cd data_collection
pip install -r requirements.txt
python collect_business_cycle_data.py
```

API Keyها اختیاری‌اند. مسیرهای keyless برای FRED، BLS-origin series و BEA public downloads در نظر گرفته شده‌اند تا build در GitHub Actions قابل‌بازتولید بماند.

## GitHub Actions

Workflow با نام `Collect Business Cycle Data` فقط هنگام تغییر ورودی‌های اجرایی data collection اجرا می‌شود و workbook را به‌عنوان Artifact ذخیره می‌کند.

Secrets اختیاری:

- `FRED_API_KEY`
- `BEA_API_KEY`
- `BLS_API_KEY`

## اصل پژوهشی

- Missing به صفر تبدیل نمی‌شود.
- interpolation برای پر کردن شکاف‌های منبع انجام نمی‌شود.
- forecast با observed قاطی نمی‌شود.
- proxy با متغیر اصلی یکی فرض نمی‌شود.
- تا قبل از دریافت specification دقیق استاد، estimator نهایی انتخاب نمی‌شود.

جزئیات آخرین اجرای اعتبارسنجی‌شده در `data_collection/RUN_STATUS.md` ثبت شده است.
