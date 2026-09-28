# AI, Digital Economy & Business Cycle Dynamics

این ریپو مربوط به پروژه اصلی بررسی اثر اقتصاد دیجیتال و هوش مصنوعی بر پویایی چرخه‌های تجاری است.

## فاز فعلی: جمع‌آوری داده

در این مرحله هنوز سراغ برآورد نهایی نمی‌رویم. اول باید داده‌های موردنیاز از منابع رسمی جمع شوند، پوشش زمانی و فرکانس آن‌ها مشخص شود و کیفیت هر سری بررسی شود.

اسکریپت فاز اول در مسیر زیر قرار دارد:

`data_collection/collect_business_cycle_data.py`

خروجی:

`AI_Digital_Economy_BusinessCycle_Data.xlsx`

## منابع فعلی

- FRED
- World Bank Open Data
- BEA
- BLS
- OECD SDMX
- IMF DataMapper
- منابع دستی: Census BTOS، World Bank DAI، Stanford HAI

## اجرای محلی

```bash
cd data_collection
pip install -r requirements.txt
python collect_business_cycle_data.py
```

برای FRED و BEA باید API Key تنظیم شود. BLS بدون کلید هم قابل استفاده است ولی محدودیت بیشتری دارد.

## GitHub Actions

Workflow با نام `Collect Business Cycle Data` هنگام تغییر فایل‌های بخش data collection اجرا می‌شود و فایل Excel را به‌عنوان Artifact ذخیره می‌کند.

برای کامل شدن خروجی، این GitHub Secrets را می‌توان اضافه کرد:

- `FRED_API_KEY`
- `BEA_API_KEY`
- `BLS_API_KEY` (اختیاری)

## مرحله بعد

بعد از اینکه خروجی داده را گرفتیم:

1. شیت Collection Status را audit می‌کنیم.
2. فرکانس ماهانه/فصلی/سالانه را تفکیک می‌کنیم.
3. Track A و Track B را بر اساس پوشش واقعی داده طراحی می‌کنیم.
4. سپس مدل اقتصادسنجی Python پیاده می‌شود.

> فایل‌ها و نمونه‌های قدیمی TFP مبنای این پروژه نیستند.
