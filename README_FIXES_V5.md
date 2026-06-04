# Fixed v4 - user request patch

Ushbu versiya `Tg-bot-main-fixed-v3` ustiga qo'shimcha tuzatishlar bilan tayyorlandi.

## Qo'shilgan va tuzatilganlar

- Ma'lumot kiritish talab qilingan flowlarda `❌ Bekor qilish` tugmasi qo'shildi.
- Telefon contact so'ralganda ham `❌ Bekor qilish` mavjud.
- Shikoyat matni yuborilgandan keyin bot qotib qolmasligi uchun FSM to'liq tozalanadi va foydalanuvchiga bosh menyu qaytariladi.
- Shikoyat/taklif admin xabarnomalari asosiy user flowini to'sib qo'ymasligi uchun xavfsiz background task orqali yuboriladi.
- `suggestions` jadvali qo'shildi.
- User menyuga `💡 Taklif` bo'limi qo'shildi.
- Admin panelga `💡 Takliflar` bo'limi qo'shildi.
- Admin takliflarni ko'rishi va statusini `Yangi`, `Ko'rib chiqilmoqda`, `Yopilgan` qilib o'zgartirishi mumkin.
- Dashboardda takliflar soni ko'rsatiladi.
- Excel exportga `Suggestions` sheet qo'shildi.
- Backup ZIP ichidagi Excel exportda takliflar ham bo'ladi.

## Tekshirilgan

- `python3 -m py_compile main.py` sintaksis tekshiruvidan o'tdi.
