# Tg-bot-main-fixed-v4

Ushbu versiyada fixed-v3 ustiga quyidagilar qo‘shildi/tuzatildi:

- Har bir ma’lumot kiritish bosqichiga `❌ Bekor qilish` tugmasi qo‘shildi.
- Shikoyat matni yuborilgandan keyin FSM to‘liq `clear()` qilinadi, reply keyboard o‘chiriladi va bosh menyu qaytariladi.
- Adminlarga yuboriladigan shikoyat/taklif matnlari HTML escape qilindi, shuning uchun foydalanuvchi `<`, `>` kabi belgilar yozsa bot qotmaydi.
- Baholashdan alohida eski ZIPdagi mantiqqa mos `🗳 Ovoz berish` tizimi qo‘shildi.
- Admin panelga `🗳 Ovozlar` bo‘limi qo‘shildi: umumiy natija, kafedralar bo‘yicha natija, ovozlarni tozalash.
- `💡 Taklif` tizimi qo‘shildi: kafedra → o‘qituvchi → taklif matni.
- Admin panelga `💡 Takliflar` bo‘limi qo‘shildi: ko‘rish va statuslarni o‘zgartirish.
- Database migrationga `suggestions`, `complaints.kind` va votes indexlari qo‘shildi.
- Excel exportga `Votes` va `Suggestions` sheetlari qo‘shildi.

Ishga tushirish:

```bash
pip install -r requirements.txt
export BOT_TOKEN="TOKEN"
python main.py
```

Railway/Vercel/Railway Variables:
- BOT_TOKEN
- ADMIN_IDS
- CHANNEL_USERNAME
- CHANNEL_URL
- DATA_DIR (ixtiyoriy)
