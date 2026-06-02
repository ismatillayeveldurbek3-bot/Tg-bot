# Tuzatilgan versiya

Ushbu versiyada oldingi modernized ZIPdagi mantiqiy va callback xatolari tuzatildi.

## Asosiy tuzatishlar
- Ishlamagan `edit_teacher` callback handleri to‘liq tuzatildi.
- O‘qituvchi ro‘yxati endi faqat matn emas, har bir o‘qituvchi uchun alohida boshqaruv oynasi ochadi.
- O‘qituvchi ismini tahrirlash, o‘quvchilar sonini o‘zgartirish, arxivlash va tiklash tugmalari ishlaydi.
- Bir foydalanuvchi bir o‘qituvchini faqat bir marta baholaydi; eski bahoni qayta o‘zgartirish bloklandi.
- Avval baholangan o‘qituvchiga kirganda baholash yulduzlari chiqmaydi, ogohlantirish chiqadi.
- Shikoyatlar admin panelda ro‘yxat ko‘rinishida ochiladi, har bir shikoyat alohida ko‘riladi.
- Shikoyat statuslari: `Yangi`, `Ko‘rib chiqilmoqda`, `Yopilgan` tugmalar orqali o‘zgartiriladi.
- Kafedra boshqaruvi qo‘shildi: qo‘shish, ro‘yxat, nomini tahrirlash, arxivlash, tiklash.
- Sozlamalardagi bo‘limlar endi real boshqaruv oynalariga ulanadi.
- Backup sozlamasida avtomatik backupni yoqish/o‘chirish qo‘shildi.
- Qo‘shimcha admin qo‘shish/olib tashlash ishlaydi.
- Teacher key yaratishda collision ehtimoli kamaytirildi.
- Import/syntax tekshiruvdan o‘tkazildi: `python3 -m py_compile main.py`.

## Ishga tushirish
Railway/VPS muhitida environment variables:

```env
BOT_TOKEN=your_bot_token
ADMIN_IDS=5298063089,7361393654
CHANNEL_USERNAME=@Qashqadaryo_PMM
CHANNEL_URL=https://t.me/Qashqadaryo_PMM
DATA_DIR=/app/data
```

So‘ng:

```bash
pip install -r requirements.txt
python main.py
```
